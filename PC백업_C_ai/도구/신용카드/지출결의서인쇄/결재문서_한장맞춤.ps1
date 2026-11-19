param([string]$In = '', [int[]]$Docs = @(), [int]$MaxH = 1300)
# 조금 넘친 결재문서를 배율을 줄여 1쪽으로. '인쇄 미리보기창을 닫으시겠습니까?' 안내창(인쇄 뒤 뜸)은 가리고 PDF
$S = $PSScriptRoot
. "$S\cdp.ps1"
$ErrorActionPreference = 'Stop'
if (-not $In) { $In = "$env:TEMP\지출결의서인쇄\모음" }
$man = @((Get-Content "$In\manifest.json" -Raw -Encoding UTF8) | ConvertFrom-Json | ForEach-Object { $_ })
function PageCount($bytes) { ([regex]::Matches([Text.Encoding]::ASCII.GetString($bytes), '/Type\s*/Page[^s]')).Count }
$hide = @'
(() => {
  let n = 0;
  const vw = innerWidth, vh = innerHeight;
  for (const e of document.querySelectorAll('body *')) {
    const cs = getComputedStyle(e); if (cs.position !== 'fixed' || cs.display === 'none') continue;
    const r = e.getBoundingClientRect();
    if ((r.width >= vw * 0.8 && r.height >= vh * 0.8) || (e.innerText || '').includes('미리보기창을 닫으시겠습니까')) { e.style.setProperty('display', 'none', 'important'); n++; }
  }
  return n;
})()
'@
$tab = New-Tab 'about:blank'
Connect-Tab $tab
$null = Cdp 'Page.enable'
$null = Cdp 'Page.addScriptToEvaluateOnNewDocument' @{ source = 'window.print = function () { window.__printed = (window.__printed || 0) + 1; };' }
foreach ($d in $man) {
  if ($Docs.Count -and $Docs -notcontains [int]$d.n) { continue }
  if (-not $Docs.Count -and (PageCount ([IO.File]::ReadAllBytes($d.pdf))) -ne 2) { continue }
  $pu = "https://gw.ds-al.com/#popup?MicroModuleCode=eap&area=2&docAuth=0&docId=$($d.docId)&docToken=&formId=175&mode=P&spApprokey=&spDocId=$($d.docId)%3B0&callComp=UBAP018&popupUUID=$([guid]::NewGuid())"
  $null = Cdp 'Page.navigate' @{ url = 'about:blank' }; Start-Sleep -Milliseconds 300
  $null = Cdp 'Page.navigate' @{ url = $pu }
  $null = WaitJs "document.body && document.body.innerText.includes('$($d.docNo)')" 40
  $null = WaitJs "(window.__printed || 0) >= 1" 30
  $null = WaitJs "[...document.images].every(i => i.complete)" 15
  Start-Sleep -Milliseconds 1500
  $h = Js 'document.documentElement.scrollHeight'
  if ($h -gt $MaxH) { '{0:00} {1} 높이 {2} → 그대로 2쪽' -f $d.n, $d.docNo, $h; continue }
  $done = $false
  foreach ($sc in 0.92, 0.88, 0.84, 0.8) {
    $hid = Js $hide
    $pdf = Cdp 'Page.printToPDF' @{ printBackground = $true; preferCSSPageSize = $true; paperWidth = 8.27; paperHeight = 11.69; marginTop = 0.4; marginBottom = 0.4; marginLeft = 0.4; marginRight = 0.4; scale = $sc } 120
    $bytes = [Convert]::FromBase64String($pdf['data'])
    if ((PageCount $bytes) -eq 1) { [IO.File]::WriteAllBytes($d.pdf, $bytes); '{0:00} {1} 높이 {2} → 배율 {3}로 1쪽 (가린 창 {4})' -f $d.n, $d.docNo, $h, $sc, $hid; $done = $true; break }
  }
  if (-not $done) { '{0:00} {1} 높이 {2} → 줄여도 2쪽' -f $d.n, $d.docNo, $h }
}
Disconnect-Tab
Close-Tab $tab.id
