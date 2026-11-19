param(
  [string]$ListFile = "$env:TEMP\지출결의서인쇄\목록.json",
  [string]$Out = "$env:TEMP\카드전표\자료.json"
)
# 지출결의서마다 지출내역(용도·내용·거래처·프로젝트·거래일자·공급가액·부가세·사용부서·사원)과
# 신용카드매출전표(가맹점·사업자번호·카드·승인번호·거래일시)를 읽어 JSON으로
$S = $PSScriptRoot
. "$S\cdp.ps1"
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
$list = @(((Get-Content $ListFile -Raw -Encoding UTF8) | ConvertFrom-Json).rows | Where-Object { $_.FORM_NM -eq '지출결의서' })
function Go($url, $cond, $sec = 40) {
  $null = Cdp 'Page.navigate' @{ url = 'about:blank' }; Start-Sleep -Milliseconds 300
  $null = Cdp 'Page.navigate' @{ url = $url }
  if (-not (WaitJs $cond $sec)) { if (Js "!!document.querySelector('input[type=password]')") { throw '작업용 엣지에서 그룹웨어 로그인이 풀렸습니다.' }; return $false }
  return $true
}
$docJs = @'
(() => {
  const txt = e => (e ? (e.innerText || '').replace(/\s+/g, ' ').trim() : '');
  // 기본정보 (라벨 다음 칸)
  const info = {};
  for (const th of document.querySelectorAll('th,td')) {
    const k = txt(th);
    if (['총합계', '작성자/작성부서', '회계단위', '프로젝트', '품 의 번 호', '기 안 부 서', '작 성 일 자'].includes(k) && th.nextElementSibling && !info[k]) info[k] = txt(th.nextElementSibling);
  }
  // 지출내역 표: '용도' '내용' '거래처' '증빙' 머리글이 있는 표 중 '거래일자'가 있는 것
  const tables = [...document.querySelectorAll('table')].filter(t => /용도/.test(t.innerText) && /거래일자/.test(t.innerText) && /증빙/.test(t.innerText));
  const t = tables[tables.length - 1];
  const items = [];
  if (t) {
    const trs = [...t.querySelectorAll('tbody tr')].filter(tr => !tr.querySelector('th'));
    for (let i = 0; i < trs.length; i++) {
      const c = [...trs[i].children];
      if (c.length >= 6 && /^\d+$/.test(txt(c[0]))) {
        const r2 = trs[i + 1] ? [...trs[i + 1].children] : []; const r3 = trs[i + 2] ? [...trs[i + 2].children] : [];
        const a = c[4].querySelector('a');
        items.push({ no: txt(c[0]), use: txt(c[1]), desc: txt(c[2]), vendor: txt(c[3]), proof: txt(c[4]), proj: txt(c[5]), link: a ? a.href : '',
          date: txt(r2[0]), due: txt(r2[1]), sup: txt(r2[2]), vat: txt(r2[3]), total: txt(r2[4]),
          bank: txt(r3[0]), acct: txt(r3[1]), holder: txt(r3[2]), dept: txt(r3[3]), emp: txt(r3[4]) });
        i += 2;
      }
    }
  }
  return JSON.stringify({ info, items });
})()
'@
$rcptJs = @'
(() => {
  const lines = document.body.innerText.split('\n').map(s => s.trim()).filter(Boolean);
  const get = k => { const i = lines.indexOf(k); return i >= 0 ? (lines[i + 1] || '') : ''; };
  const amt = lines.find(s => /^[\d,]+\s*원$/.test(s) || /^[\d,]+$/.test(s) && false) || '';
  return JSON.stringify({ name: lines[0] || '', amount: (lines.find(s => /원$/.test(s)) || '').replace(/[^\d]/g, ''), bizNo: get('사업자등록번호'), ceo: get('대표자명'), tel: get('전화번호'), addr: get('주소'), kind: get('업종'),
    cardAlias: get('카드별칭'), cardNo: get('카드번호'), sup: get('공급가액'), vat: get('부가세'), svc: get('봉사료'), approval: get('승인번호'), dt: get('거래일시'), raw: lines.slice(0, 40).join(' | ') });
})()
'@
$tab = New-Tab 'about:blank'
Connect-Tab $tab
$null = Cdp 'Page.enable'
$docs = New-Object Collections.ArrayList; $n = 0
foreach ($d in $list) {
  $n++
  $du = "https://gw.ds-al.com/#/popup?MicroModuleCode=eap&docAuth=0&docID=$($d.DOC_ID)&formId=$($d.FORM_ID)&pageCode=UBA1100&callComp=UBAP002&popupUUID=$([guid]::NewGuid())"
  if (-not (Go $du "document.body && document.body.innerText.includes('$($d.DOC_NO)') && document.body.innerText.includes('지출내역')" 40)) { "$n 문서 안 열림: $($d.DOC_NO)"; continue }
  Start-Sleep -Milliseconds 800
  $dj = $script:Ser.DeserializeObject((Js $docJs))
  $items = @($dj['items'])
  foreach ($it in $items) {
    if ($it['link']) {
      $null = Cdp 'Emulation.setDeviceMetricsOverride' @{ width = 380; height = 900; deviceScaleFactor = 1; mobile = $false }
      try {
        if (Go $it['link'] "document.body && document.body.innerText.includes('승인번호') && document.body.innerText.includes('거래일시')" 30) { Start-Sleep -Milliseconds 400; $it['receipt'] = $script:Ser.DeserializeObject((Js $rcptJs)) }
        else { $it['receipt'] = $null }
      } finally { $null = Cdp 'Emulation.clearDeviceMetricsOverride' }
    }
  }
  [void]$docs.Add(@{ n = $n; docId = $d.DOC_ID; docNo = $d.DOC_NO; title = $d.DOC_TITLE; info = $dj['info']; items = $items })
  "{0:00} {1} {2}: 지출내역 {3}줄, 영수증 {4}장" -f $n, $d.DOC_NO, $d.DOC_TITLE, $items.Count, @($items | Where-Object { $_['receipt'] }).Count
}
Disconnect-Tab
Close-Tab $tab.id
[IO.File]::WriteAllText($Out, $script:Ser.Serialize($docs), (New-Object Text.UTF8Encoding $false))
"저장: $Out"
