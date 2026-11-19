param(
  [string]$ListFile = "$env:TEMP\지출결의서인쇄\목록.json",   # 목록_읽기.ps1 결과 (지출결의서만 씀)
  [string]$OutDir = "$env:USERPROFILE\Desktop",
  [int]$Scale = 2,
  [switch]$Print                                            # 엑셀로 만든 뒤 기본 프린터로 인쇄
)
# 지출결의서 문서들의 파란 글씨 '신용카드매출전표' 링크 → 영수증 부분만 그림 → 엑셀(영수증 시트)에 차례로 넣고 영수증_8장정리(A4 한 장 8장) → (인쇄)
$S = $PSScriptRoot
. "$S\cdp.ps1"
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing, System.Runtime.WindowsRuntime
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
$asTaskA = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } | Select-Object -First 1
function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(); $task.Result }
function AwaitA($op) { $asTaskA.Invoke($null, @($op)).Wait() }
function PdfPage1($path, $width) {
  $f = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($path)) ([Windows.Storage.StorageFile])
  $doc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($f)) ([Windows.Data.Pdf.PdfDocument])
  $pg = $doc.GetPage(0); $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions; $opt.DestinationWidth = $width
  $ms = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream; AwaitA ($pg.RenderToStreamAsync($ms, $opt))
  $b = New-Object System.Drawing.Bitmap ([System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($ms.GetInputStreamAt(0)))
  $copy = New-Object System.Drawing.Bitmap $b; $b.Dispose(); $pg.Dispose(); return $copy
}
function Go($url, $cond, $sec = 40) {
  $null = Cdp 'Page.navigate' @{ url = 'about:blank' }; Start-Sleep -Milliseconds 300
  $null = Cdp 'Page.navigate' @{ url = $url }
  if (-not (WaitJs $cond $sec)) { if (Js "!!document.querySelector('input[type=password]')") { throw '작업용 엣지에서 그룹웨어 로그인이 풀렸습니다.' }; return $false }
  return $true
}
$findBox = @'
(() => {
  const has = (e, ws) => ws.every(w => (e.innerText || '').includes(w));
  const all = [...document.querySelectorAll('body *')].filter(e => e.offsetParent !== null || e.tagName === 'BODY');
  let c = all.filter(e => has(e, ['승인번호', '거래일시']));
  if (!c.length) return null;
  let b = c[c.length - 1];
  while (b.parentElement && !/\d[\d,]*\s*원/.test(b.innerText)) b = b.parentElement;
  if (b.parentElement && b.parentElement.innerText.length < b.innerText.length * 1.4) b = b.parentElement;
  let L = 1e9, T = 1e9, R = -1e9, B = -1e9;
  [b, ...b.querySelectorAll('*')].forEach(e => {
    const leaf = !e.children.length || e.tagName === 'svg' || e.tagName === 'IMG';
    const cs = getComputedStyle(e); const line = cs.borderBottomStyle !== 'none' || cs.borderTopStyle !== 'none';
    if (!leaf && !line) return;
    if (leaf && e.tagName !== 'svg' && e.tagName !== 'IMG' && !(e.innerText || '').trim() && !line) return;
    const q = e.getBoundingClientRect(); if (!q.width || !q.height) return;
    L = Math.min(L, q.left); T = Math.min(T, q.top); R = Math.max(R, q.right); B = Math.max(B, q.bottom);
  });
  const r0 = b.getBoundingClientRect(); if (R < L) { L = r0.left; T = r0.top; R = r0.right; B = r0.bottom; }
  return { x: L + scrollX, y: T + scrollY, w: R - L, h: B - T, docH: document.documentElement.scrollHeight };
})()
'@

$list = @(((Get-Content $ListFile -Raw -Encoding UTF8) | ConvertFrom-Json).rows | Where-Object { $_.FORM_NM -eq '지출결의서' })
$tmp = "$env:TEMP\지출결의서인쇄\영수증"; New-Item -ItemType Directory -Force $tmp | Out-Null
"지출결의서 $($list.Count)건에서 영수증 모으기"
$tab = New-Tab 'about:blank'
Connect-Tab $tab
$null = Cdp 'Page.enable'
$shots = New-Object Collections.Generic.List[string]; $n = 0
foreach ($d in $list) {
  $n++; $nn = '{0:00}' -f $n
  $du = "https://gw.ds-al.com/#/popup?MicroModuleCode=eap&docAuth=0&docID=$($d.DOC_ID)&formId=$($d.FORM_ID)&pageCode=UBA1100&callComp=UBAP002&popupUUID=$([guid]::NewGuid())"
  if (-not (Go $du "document.body && document.body.innerText.includes('$($d.DOC_NO)') && document.body.innerText.includes('지출내역')" 40)) { "$nn $($d.DOC_NO): 문서가 안 열림"; continue }
  Start-Sleep -Milliseconds 800
  $links = @(Js "[...new Set([...document.querySelectorAll('a')].filter(a => a.getAttribute('href') && ((a.innerText||'').includes('매출전표') || a.href.includes('APB1020'))).map(a => a.href))]")
  $r = 0
  foreach ($u in $links) {
    $r++
    $null = Cdp 'Emulation.setDeviceMetricsOverride' @{ width = 380; height = 900; deviceScaleFactor = 1; mobile = $false }
    try {
      if (-not (Go $u "document.body && document.body.innerText.includes('승인번호') && document.body.innerText.includes('거래일시')" 30)) { "  영수증 $r 안 열림"; continue }
      Start-Sleep -Milliseconds 600
      $bx = Js $findBox; if (-not $bx) { "  영수증 $r 상자 못 찾음"; continue }
      $paperH = [Math]::Max(($bx['docH'] + 10), ($bx['y'] + $bx['h'] + 20)) / 96
      $rp = Cdp 'Page.printToPDF' @{ printBackground = $true; paperWidth = (380 / 96); paperHeight = $paperH; marginTop = 0; marginBottom = 0; marginLeft = 0; marginRight = 0; pageRanges = '1' } 60
      $pf = "$tmp\_tmp.pdf"; [IO.File]::WriteAllBytes($pf, [Convert]::FromBase64String($rp['data']))
      $bmp = PdfPage1 $pf (380 * $Scale)
      $x = [Math]::Max(0, [int](($bx['x'] - 6) * $Scale)); $y = [Math]::Max(0, [int](($bx['y'] - 6) * $Scale))
      $w = [Math]::Min($bmp.Width - $x, [int](($bx['w'] + 12) * $Scale)); $h = [Math]::Min($bmp.Height - $y, [int](($bx['h'] + 12) * $Scale))
      $crop = $bmp.Clone((New-Object System.Drawing.Rectangle $x, $y, $w, $h), $bmp.PixelFormat)
      $p = "$tmp\R_$($nn)_$('{0:00}' -f $r).png"; $crop.Save($p, [System.Drawing.Imaging.ImageFormat]::Png); $crop.Dispose(); $bmp.Dispose(); Remove-Item -LiteralPath $pf -Force
      $shots.Add($p)
    } finally { $null = Cdp 'Emulation.clearDeviceMetricsOverride' }
  }
  "$nn $($d.DOC_NO) $($d.DOC_TITLE): 영수증 $r 장"
}
Disconnect-Tab
Close-Tab $tab.id
if (-not $shots.Count) { throw '모은 영수증이 없습니다.' }

# ── 엑셀에 넣고 8장씩 정리 (영수증_링크로_모으기와 같은 방식) ──
$stamp = Get-Date -Format 'yyyyMMdd_HHmm'
$out = Join-Path $OutDir "영수증_8장_$stamp.xlsx"
$xl = $null; for ($k = 0; $k -lt 6 -and -not $xl; $k++) { try { $xl = New-Object -ComObject Excel.Application } catch { Start-Sleep 3 } }
if (-not $xl) { throw '엑셀을 열 수 없습니다.' }
$xl.DisplayAlerts = $false; $wb = $xl.Workbooks.Add(); $sh = $wb.Worksheets.Item(1); $sh.Name = '영수증'
$i = 0; foreach ($f in $shots) { $sh.Shapes.AddPicture($f, 0, -1, 10, 10 + $i * 50, -1, -1) | Out-Null; $i++ }
$wb.SaveAs($out, 51); $wb.Close(); $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl); $xl = $null
& (Join-Path (Split-Path $S) '영수증_8장정리.ps1') -Workbook $out | Out-Null
"엑셀: $out (영수증 $($shots.Count)장 → A4 $([Math]::Ceiling($shots.Count / 8))장)"
if ($Print) {
  $xl = New-Object -ComObject Excel.Application; $xl.DisplayAlerts = $false
  $wb = $xl.Workbooks.Open($out, 0, $true)
  $pages = $wb.Worksheets.Item(1).PageSetup.Pages.Count
  $wb.Worksheets.Item(1).PrintOut()
  $wb.Close($false); $xl.Quit()
  "인쇄 보냄: $pages 쪽 → 기본 프린터"
}
Get-ChildItem -LiteralPath $tmp -Filter 'R_*.png' | Remove-Item -Force   # 영수증 그림(임시) 지움
