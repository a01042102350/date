param([string]$Out = (Join-Path $PSScriptRoot '부서코드.csv'))
# 아마란스 부서 코드표(전체) → 부서코드.csv (부서코드, 부서명). 카드 자동전표의 '사용부서'(CT_DEPT)에 지출결의서 작성부서 코드를 넣을 때 씀.
# 작업용 엣지(포트 9333, 그룹웨어 로그인 상태)에 새 탭: 회계관리 > 자동전표처리 > 외부연동 > 자동전표처리 화면의 '작성부서' 부서코드도움 창을 열어
# 그 창이 보내는 부서 목록 요청(CP0001A0001, DEPT_SHOWGRP_CODE)의 회계단위 조건을 빼서 전체를 받음. 읽기만 함. (2026-11-27 만듦: 85곳)
. (Join-Path $PSScriptRoot 'cdp.ps1')
$ErrorActionPreference = 'Stop'
$tab = New-Tab 'https://gw.ds-al.com/#/'
Connect-Tab $tab
function ClickMenu($txt) {
  Js "(() => { const c = Array.from(document.querySelectorAll('*')).filter(e => e.children.length === 0 && (e.innerText || '').trim() === '$txt' && e.offsetParent !== null); const e = c.pop(); if (!e) return 'none'; ['mousedown','mouseup','click'].forEach(t => e.dispatchEvent(new MouseEvent(t, {bubbles: true}))); return 'ok'; })()"
}
try {
  if (-not (WaitJs "document.body && document.body.innerText.indexOf('회계관리') >= 0" 40)) { throw '그룹웨어 첫 화면이 안 열립니다(로그인 확인)' }
  Start-Sleep -Seconds 2
  [void](Js @'
(() => { window.__dept = null;
  const o = XMLHttpRequest.prototype.open, s = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function(m, u) { this.__u = u; return o.apply(this, arguments); };
  XMLHttpRequest.prototype.send = function(b) {
    if (/CP0001A0001$/.test(this.__u) && typeof b === 'string') {
      try { const j = JSON.parse(b); if (j.helpTy === 'DEPT_SHOWGRP_CODE') { j.additionalConditions = {}; b = JSON.stringify(j); this.addEventListener('load', () => { window.__dept = this.responseText || 'empty'; }); } } catch (e) {}
    }
    return s.call(this, b); };
  return 'ok'; })()
'@)
  foreach ($m in '회계관리', '자동전표처리', '외부연동', '자동전표처리') { $r = ClickMenu $m; if ($r -ne 'ok') { throw "메뉴 '$m'을 못 찾았습니다" }; Start-Sleep -Seconds 3 }
  Start-Sleep -Seconds 3
  $r = Js @'
(() => { const inp = Array.from(document.querySelectorAll('input')).filter(i => (i.placeholder || '').indexOf('부서코드도움') >= 0 && i.offsetParent !== null)[0];
  if (!inp) return 'no input'; let box = inp.parentElement, btn = null;
  for (let i = 0; i < 4 && box && !btn; i++) { btn = box.querySelector('button'); box = box.parentElement; }
  if (!btn) return 'no button'; ['mousedown','mouseup','click'].forEach(t => btn.dispatchEvent(new MouseEvent(t, {bubbles: true}))); return 'ok'; })()
'@
  if ($r -ne 'ok') { throw "작성부서 코드도움을 못 열었습니다($r)" }
  if (-not (WaitJs '!!window.__dept' 30)) { throw '부서 목록 응답이 없습니다' }
  $o = (Js 'window.__dept' 60) | ConvertFrom-Json
  if ($o.resultCode -ne 0) { throw "부서 목록 오류: $($o.resultMsg)" }
  $rows = @($o.resultData | ForEach-Object { [pscustomobject]@{ 부서코드 = "$($_.deptCd)"; 부서명 = "$($_.deptNm)".Trim() } })
  if ($rows.Count -lt 10) { throw "부서가 $($rows.Count)곳뿐이라 이상합니다" }
  $rows | Export-Csv $Out -NoTypeInformation -Encoding UTF8
  "아마란스 부서 $($rows.Count)곳 → $Out"
} finally { [void](Js "document.dispatchEvent(new KeyboardEvent('keydown', {key: 'Escape', keyCode: 27, bubbles: true}))"); Disconnect-Tab; Close-Tab $tab.id }
