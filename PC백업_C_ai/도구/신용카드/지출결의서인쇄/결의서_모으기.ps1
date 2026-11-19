param(
  [string]$Out = '',
  [int]$MaxAttachPages = 12,
  [string]$Only = ''            # 시험용: 문서번호 일부(예: 12012)만
)
# 미결문서 목록(목록.json)의 지출결의서를 차례로: 결재문서 PDF, 첨부(미리보기 쪽 그림), 신용카드매출전표(영수증 부분만) 모으기
$S = $PSScriptRoot
. "$S\cdp.ps1"
$ErrorActionPreference = 'Stop'
if (-not $Out) { $Out = "$env:TEMP\지출결의서인쇄\모음" }
New-Item -ItemType Directory -Force $Out | Out-Null

# ── PDF → 그림 (WinRT) ──
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
function SaveDataUrl($dataUrl, $pathNoExt) {
  $i = $dataUrl.IndexOf(';base64,'); if (-not $dataUrl.StartsWith('data:') -or $i -lt 0) { throw 'data url 아님' }
  $mime = $dataUrl.Substring(5, $i - 5); $b64 = $dataUrl.Substring($i + 8)   # switch -Regex가 $Matches를 덮어써서 0바이트가 되던 문제
  $ext = if ($mime -match 'jpe?g') { 'jpg' } elseif ($mime -match 'gif') { 'gif' } elseif ($mime -match 'bmp') { 'bmp' } else { 'png' }
  $p = "$pathNoExt.$ext"; [IO.File]::WriteAllBytes($p, [Convert]::FromBase64String($b64)); return $p
}
function Go($url, $cond, $sec = 40) {
  $null = Cdp 'Page.navigate' @{ url = 'about:blank' }; Start-Sleep -Milliseconds 300
  $null = Cdp 'Page.navigate' @{ url = $url }
  if (-not (WaitJs $cond $sec)) {
    if (Js "!!document.querySelector('input[type=password]')") { throw '작업용 엣지에서 그룹웨어 로그인이 풀렸습니다.' }
    return $false
  }
  return $true
}

$list = @(((Get-Content "$env:TEMP\지출결의서인쇄\목록.json" -Raw -Encoding UTF8) | ConvertFrom-Json).rows | Where-Object { $_.FORM_NM -eq '지출결의서' })
if ($Only) { $list = @($list | Where-Object { $_.DOC_NO -like "*$Only*" }) }
"지출결의서 $($list.Count)건"

$tab = New-Tab 'about:blank'
Connect-Tab $tab
$null = Cdp 'Page.enable'
$null = Cdp 'Page.addScriptToEvaluateOnNewDocument' @{ source = 'window.print = function () { window.__printed = (window.__printed || 0) + 1; };' }

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
  return { x: L + scrollX, y: T + scrollY, w: R - L, h: B - T, docH: document.documentElement.scrollHeight, text: b.innerText.slice(0, 60) };
})()
'@

$docInfo = @'
(() => {
  const links = [...document.querySelectorAll('a')].filter(a => (a.innerText || '').includes('매출전표') && a.getAttribute('href')).map(a => a.href);
  const boxes = [...document.querySelectorAll('div.UpDownLoader')];
  let files = [], auth = null;
  for (const box of boxes) {
    const fk = Object.keys(box).find(k => k.startsWith('__reactFiber')); if (!fk) continue;
    let f = box[fk]; for (let i = 0; i < 6 && f; i++, f = f.return) {
      const c = f.stateNode;
      if (c && c.state && Array.isArray(c.state.bindinglist) && c.props && String(c.props.moduleGbn).toUpperCase() === 'EAP') {
        files = c.state.bindinglist.map(x => ({ sn: x.fileSn, name: x.fileName, ext: x.fileExtsn, size: x.fileSize })); auth = c.props.authKeyMap; break;
      }
    }
    if (auth) break;
  }
  const t = document.body.innerText;
  const m = t.match(/첨부파일\s*(\d+)개/);
  return JSON.stringify({ links: [...new Set(links)], files, auth, attachHead: m ? Number(m[1]) : 0, title: document.title });
})()
'@

$viewerPages = @'
(async () => {
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const total = Number((document.getElementById('total-page') || {}).innerText || 0);
  if (total > __MAX__) return JSON.stringify({ total, pages: [], tooMany: true });
  let imgs = [];
  const ok = i => (i.src.startsWith('blob:') || i.src.includes('/SynapDocViewServer/result/')) && i.complete && i.naturalWidth > 1;
  for (let t = 0; t < 80; t++) {
    imgs = [...document.querySelectorAll('img.contents-page__img')];
    for (const i of imgs) if (!ok(i)) { i.scrollIntoView(); break; }
    const ready = imgs.filter(ok);
    if (total && ready.length >= total) break;
    await sleep(500);
  }
  const ready = imgs.filter(ok);
  const pages = [];
  for (const i of ready) {
    const b = await (await fetch(i.src)).blob();
    pages.push(await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(b); }));
  }
  return JSON.stringify({ total, pages, sizes: ready.map(i => i.naturalWidth + 'x' + i.naturalHeight) });
})()
'@.Replace('__MAX__', [string]$MaxAttachPages)

$hideDlg = @'
(() => {
  for (const e of document.querySelectorAll('body *')) {
    const cs = getComputedStyle(e); if (cs.position !== 'fixed' || cs.display === 'none') continue;
    const r = e.getBoundingClientRect();
    if ((r.width >= innerWidth * 0.8 && r.height >= innerHeight * 0.8) || (e.innerText || '').includes('미리보기창을 닫으시겠습니까')) e.style.setProperty('display', 'none', 'important');
  }
  return 1;
})()
'@
function PdfPageCount($bytes) { ([regex]::Matches([Text.Encoding]::ASCII.GetString($bytes), '/Type\s*/Page[^s]')).Count }

$manifest = New-Object Collections.ArrayList
$n = 0
foreach ($d in $list) {
  $n++; $nn = '{0:00}' -f $n
  $item = [ordered]@{ n = $n; docId = $d.DOC_ID; docNo = $d.DOC_NO; title = $d.DOC_TITLE; user = $d.USER_NM; pdf = $null; attach = @(); receipts = @(); notes = @() }
  try {
    # 1) 결재문서 (인쇄 화면)
    $pu = "https://gw.ds-al.com/#popup?MicroModuleCode=eap&area=2&docAuth=0&docId=$($d.DOC_ID)&docToken=&formId=$($d.FORM_ID)&mode=P&spApprokey=&spDocId=$($d.DOC_ID)%3B0&callComp=UBAP018&popupUUID=$([guid]::NewGuid())"
    if (-not (Go $pu "document.body && document.body.innerText.includes('지출결의서') && document.body.innerText.includes('$($d.DOC_NO)')" 40)) { throw '결재문서 인쇄 화면이 안 열림' }
    # 인쇄 화면은 준비가 끝나면 스스로 print()를 부름 → 그때까지 기다려야 아래가 안 잘림(2026-11-13)
    $null = WaitJs "(window.__printed || 0) >= 1" 30
    $null = WaitJs "[...document.images].every(i => i.complete)" 15
    Start-Sleep -Milliseconds 1500
    $docH = Js 'document.documentElement.scrollHeight'
    $best = $null
    foreach ($sc in 1, 0.92, 0.88, 0.84) {
      if ($sc -lt 1 -and $docH -gt 1300) { break }        # 많이 넘치면 줄이지 않고 2쪽 이상 그대로
      $null = Js $hideDlg                                  # '인쇄 미리보기창을 닫으시겠습니까?' 창 가림
      $pdf = Cdp 'Page.printToPDF' @{ printBackground = $true; preferCSSPageSize = $true; paperWidth = 8.27; paperHeight = 11.69; marginTop = 0.4; marginBottom = 0.4; marginLeft = 0.4; marginRight = 0.4; scale = $sc } 120
      $bytes = [Convert]::FromBase64String($pdf['data'])
      if (-not $best) { $best = $bytes }
      if ((PdfPageCount $bytes) -eq 1) { $best = $bytes; break }
    }
    $item.pdf = "$Out\$($nn)_결재문서.pdf"; [IO.File]::WriteAllBytes($item.pdf, $best)

    # 2) 문서 화면: 매출전표 링크, 첨부 목록
    $du = "https://gw.ds-al.com/#/popup?MicroModuleCode=eap&docAuth=0&docID=$($d.DOC_ID)&formId=$($d.FORM_ID)&pageCode=UBA1100&callComp=UBAP002&popupUUID=$([guid]::NewGuid())"
    $cond = if ($d.FILE_CNT -gt 0) { "[...document.querySelectorAll('div.UpDownLoader')].some(b => /첨부파일\s*[1-9]\d*개/.test(b.innerText||''))" } else { "document.body && document.body.innerText.includes('$($d.DOC_NO)') && document.body.innerText.includes('지출내역')" }
    if (-not (Go $du $cond 40)) { $item.notes += '문서 화면이 늦게 열려 첨부·영수증을 못 읽었을 수 있음' }
    Start-Sleep -Milliseconds 800
    $di = (Js $docInfo) | ConvertFrom-Json
    $links = @($di.links)
    if ($d.FILE_CNT -gt 0 -and -not $di.files.Count) { $item.notes += "첨부 $($d.FILE_CNT)개가 있는데 목록을 못 읽음" }

    # 3) 첨부 (교육 신청서는 뺌 — 2026-11-13 사용자 지시)
    $k = 0
    foreach ($f in @($di.files)) {
      $k++; $fname = "$($f.name).$($f.ext)"
      if ($f.name -match '신청서' -and $f.name -match '교육') { $item.attach += [ordered]@{ name = $fname; pages = @(); skip = '교육 신청서라 뺌' }; continue }
      $authStr = ($di.auth | ConvertTo-Json -Compress)
      $q = 'authKeyMap=' + [Uri]::EscapeDataString(($authStr | ConvertTo-Json -Compress)) + "&fileSn=$($f.sn)&moduleGbn=EAP&path=%2Fecm%2Fecm001A07&pathSeq=Ea&callComp=UAA0700&popupUUID=$([guid]::NewGuid())"
      if (-not (Go "https://gw.ds-al.com/#popup?$q" "!!document.getElementById('total-page') && Number(document.getElementById('total-page').innerText) > 0" 90)) { $item.attach += [ordered]@{ name = $fname; pages = @(); skip = '미리보기가 열리지 않음' }; continue }
      Start-Sleep -Milliseconds 800
      $vp = $script:Ser.DeserializeObject((Js $viewerPages 180))   # 큰 글자(그림)는 ConvertFrom-Json이 비워 버림
      $vTotal = [int]$vp['total']
      if ($vp.ContainsKey('tooMany') -and $vp['tooMany']) { $item.attach += [ordered]@{ name = $fname; pages = @(); skip = "$($vTotal)쪽이라 자동 인쇄에서 뺌" }; continue }
      $paths = @(); $p = 0
      foreach ($du2 in @($vp['pages'])) { $p++; $paths += SaveDataUrl $du2 "$Out\$($nn)_첨부$($k)_$('{0:00}' -f $p)" }
      $skip = if ($paths.Count -lt $vTotal) { "전체 $($vTotal)쪽 중 $($paths.Count)쪽만 받음" } else { '' }
      $item.attach += [ordered]@{ name = $fname; pages = $paths; skip = $skip }
    }

    # 4) 신용카드매출전표 (영수증 부분만, 폭 380 화면)
    $r = 0
    foreach ($u in $links) {
      $r++
      $null = Cdp 'Emulation.setDeviceMetricsOverride' @{ width = 380; height = 900; deviceScaleFactor = 1; mobile = $false }
      try {
        if (-not (Go $u "document.body && document.body.innerText.includes('승인번호') && document.body.innerText.includes('거래일시')" 30)) { $item.notes += "영수증 $r 번째가 안 열림"; continue }
        Start-Sleep -Milliseconds 700
        $bx = Js $findBox
        if (-not $bx) { $item.notes += "영수증 $r 번째 상자를 못 찾음"; continue }
        $paperH = [Math]::Max(($bx['docH'] + 10), ($bx['y'] + $bx['h'] + 20)) / 96
        $rp = Cdp 'Page.printToPDF' @{ printBackground = $true; paperWidth = (380 / 96); paperHeight = $paperH; marginTop = 0; marginBottom = 0; marginLeft = 0; marginRight = 0; pageRanges = '1' } 60
        $tmp = "$Out\_rcpt_tmp.pdf"; [IO.File]::WriteAllBytes($tmp, [Convert]::FromBase64String($rp['data']))
        $sc = 2; $bmp = PdfPage1 $tmp (380 * $sc)
        $x = [Math]::Max(0, [int](($bx['x'] - 6) * $sc)); $y = [Math]::Max(0, [int](($bx['y'] - 6) * $sc))
        $w = [Math]::Min($bmp.Width - $x, [int](($bx['w'] + 12) * $sc)); $h = [Math]::Min($bmp.Height - $y, [int](($bx['h'] + 12) * $sc))
        $crop = $bmp.Clone((New-Object System.Drawing.Rectangle $x, $y, $w, $h), $bmp.PixelFormat)
        $rpath = "$Out\R_$($nn)_$('{0:00}' -f $r).png"; $crop.Save($rpath, [System.Drawing.Imaging.ImageFormat]::Png); $crop.Dispose(); $bmp.Dispose(); Remove-Item $tmp -Force
        $item.receipts += $rpath
      } finally { $null = Cdp 'Emulation.clearDeviceMetricsOverride' }
    }
    if (-not $links.Count) { $item.notes += '신용카드매출전표 링크 없음' }
  } catch { $item.notes += "오류: $($_.Exception.Message)" }
  [void]$manifest.Add($item)
  $att = (@($item.attach | ForEach-Object { if ($_.skip -and -not $_.pages.Count) { "$($_.name)(뺌)" } else { "$($_.name) $($_.pages.Count)쪽" } }) -join ', ')
  "$nn $($d.DOC_NO) $($d.DOC_TITLE) | 영수증 $($item.receipts.Count) | 첨부: $att | $($item.notes -join '; ')"
}
Disconnect-Tab
Close-Tab $tab.id
[IO.File]::WriteAllText("$Out\manifest.json", ($manifest | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false))
"끝: $Out\manifest.json"
