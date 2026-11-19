param([string]$In = '')
# manifest의 결재문서 PDF를 다시 만듦: 인쇄 화면이 스스로 print()를 부를 때까지(=준비 끝) 기다리고, 안쪽 스크롤 영역을 펼친 뒤 PDF
$S = $PSScriptRoot
. "$S\cdp.ps1"
$ErrorActionPreference = 'Stop'
if (-not $In) { $In = "$env:TEMP\지출결의서인쇄\모음" }
$man = @((Get-Content "$In\manifest.json" -Raw -Encoding UTF8) | ConvertFrom-Json | ForEach-Object { $_ })   # PS 5.1은 JSON 배열을 한 덩어리로 줌
$expand = @'
(() => {
  const isScroll = e => { const cs = getComputedStyle(e); return e.scrollHeight > e.clientHeight + 5 && /(auto|scroll|hidden)/.test(cs.overflowY) && e.clientHeight > 0; };
  const sc = [...document.querySelectorAll('*')].filter(isScroll).filter(e => e !== document.documentElement && e !== document.body || true);
  const touched = new Set();
  for (const e of sc) { let n = e; while (n) { if (!touched.has(n)) { touched.add(n); const pos = getComputedStyle(n).position; n.style.setProperty('height', 'auto', 'important'); n.style.setProperty('max-height', 'none', 'important'); n.style.setProperty('overflow', 'visible', 'important'); n.style.setProperty('flex', 'none', 'important'); if (pos === 'absolute' || pos === 'fixed') n.style.setProperty('position', 'static', 'important'); } n = n.parentElement; } }
  const left = [...document.querySelectorAll('*')].filter(e => e.scrollHeight > e.clientHeight + 5 && /(auto|scroll)/.test(getComputedStyle(e).overflowY) && e.clientHeight > 0).length;
  return JSON.stringify({ touched: touched.size, left, h: document.documentElement.scrollHeight });
})()
'@
$tab = New-Tab 'about:blank'
Connect-Tab $tab
$null = Cdp 'Page.enable'
$null = Cdp 'Page.addScriptToEvaluateOnNewDocument' @{ source = 'window.print = function () { window.__printed = (window.__printed || 0) + 1; };' }
foreach ($d in $man) {
  $pu = "https://gw.ds-al.com/#popup?MicroModuleCode=eap&area=2&docAuth=0&docId=$($d.docId)&docToken=&formId=175&mode=P&spApprokey=&spDocId=$($d.docId)%3B0&callComp=UBAP018&popupUUID=$([guid]::NewGuid())"
  $null = Cdp 'Page.navigate' @{ url = 'about:blank' }; Start-Sleep -Milliseconds 300
  $null = Cdp 'Page.navigate' @{ url = $pu }
  $ok = WaitJs "document.body && document.body.innerText.includes('$($d.docNo)')" 40
  $pr = WaitJs "(window.__printed || 0) >= 1" 30
  $null = WaitJs "[...document.images].every(i => i.complete)" 15
  Start-Sleep -Milliseconds 1500
  $ex = (Js $expand) | ConvertFrom-Json
  # print() 뒤에 뜨는 '인쇄 미리보기창을 닫으시겠습니까?' 창과 회색 막을 가림
  $null = Js "(() => { for (const e of document.querySelectorAll('body *')) { const cs = getComputedStyle(e); if (cs.position !== 'fixed' || cs.display === 'none') continue; const r = e.getBoundingClientRect(); if ((r.width >= innerWidth * 0.8 && r.height >= innerHeight * 0.8) || (e.innerText || '').includes('미리보기창을 닫으시겠습니까')) e.style.setProperty('display', 'none', 'important'); } return 1; })()"
  Start-Sleep -Milliseconds 600
  $pdf = Cdp 'Page.printToPDF' @{ printBackground = $true; preferCSSPageSize = $true; paperWidth = 8.27; paperHeight = 11.69; marginTop = 0.4; marginBottom = 0.4; marginLeft = 0.4; marginRight = 0.4 } 120
  $bytes = [Convert]::FromBase64String($pdf['data'])
  [IO.File]::WriteAllBytes($d.pdf, $bytes)
  $pages = ([regex]::Matches([Text.Encoding]::ASCII.GetString($bytes), '/Type\s*/Page[^s]')).Count
  '{0:00} {1} | 열림 {2} 준비 {3} | 펼침 {4} 남은스크롤 {5} 높이 {6} | {7}쪽' -f $d.n, $d.docNo, $ok, $pr, $ex.touched, $ex.left, $ex.h, $pages
}
Disconnect-Tab
Close-Tab $tab.id
