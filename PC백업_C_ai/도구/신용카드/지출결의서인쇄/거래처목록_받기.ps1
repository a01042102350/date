param([string]$Out = "$env:TEMP\카드전표\거래처목록.json")
# 아마란스 일반거래처 전체(코드·이름·사업자번호·등록일) → JSON. 작업용 엣지(포트 9333, 그룹웨어 로그인 상태)에 새 탭을 열어 읽고 닫음. 읽기만 함.
# 방법: 회계관리 > 회계기초정보관리 > 거래처/계좌/카드관리 > 일반거래처등록 화면이 스스로 보내는 목록 요청(SYB0010_selectTradeList)의 '한 번에 100개'를 20000개로 바꿔 한 번에 받음
#  (주소 직접 입력은 첫 화면으로 튕겨서 왼쪽 메뉴를 눌러 들어감. 2026-11-26 만듦: 새 거래처 코드를 겹치지 않게 매기고, 이미 있는 거래처를 사업자번호로 찾으려고)
. (Join-Path $PSScriptRoot 'cdp.ps1')
$ErrorActionPreference = 'Stop'
$tab = New-Tab 'https://gw.ds-al.com/#/'
Connect-Tab $tab
function ClickMenu($txt) {
  Js @"
(() => { const c = Array.from(document.querySelectorAll('*')).filter(e => e.children.length === 0 && (e.innerText || '').trim() === '$txt' && e.offsetParent !== null);
  const e = c.pop(); if (!e) return 'none'; ['mousedown','mouseup','click'].forEach(t => e.dispatchEvent(new MouseEvent(t, {bubbles: true}))); return 'ok'; })()
"@
}
try {
  if (-not (WaitJs "document.body && document.body.innerText.indexOf('회계관리') >= 0" 40)) { throw '그룹웨어 첫 화면이 안 열립니다(로그인 확인)' }
  Start-Sleep -Seconds 2
  [void](Js @'
(() => { window.__trade = null;
  const o = XMLHttpRequest.prototype.open, s = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function(m, u) { this.__u = u; return o.apply(this, arguments); };
  XMLHttpRequest.prototype.send = function(b) {
    if (/SYB0010_selectTradeList$/.test(this.__u) && !window.__trade) {
      try { const j = JSON.parse(b); j.limit = 20000; j.offset = 0; b = JSON.stringify(j); } catch (e) {}
      this.addEventListener('load', () => { window.__trade = this.responseText || 'empty'; });
    }
    return s.call(this, b); };
  return 'ok'; })()
'@)
  foreach ($m in '회계관리', '회계기초정보관리', '거래처/계좌/카드관리', '일반거래처등록') {
    $r = ClickMenu $m; if ($r -ne 'ok') { throw "메뉴 '$m'을 못 찾았습니다(화면이 바뀌었을 수 있음)" }; Start-Sleep -Seconds 3
  }
  if (-not (WaitJs '!!window.__trade' 60)) { throw '거래처 목록 응답이 없습니다' }
  $json = Js @'
(() => { const j = JSON.parse(window.__trade); if (j.resultCode !== 0) return JSON.stringify({ err: 'resultCode ' + j.resultCode + ' ' + (j.resultMsg || '') });
  return JSON.stringify({ rows: (j.resultData || []).map(r => ({ trCd: r.trCd, trNm: r.trNm, regNb: r.regNb, trFg: r.trFg, insertDt: r.insertDt })) }); })()
'@ 120
  $o = $json | ConvertFrom-Json
  if ($o.err) { throw "거래처 목록 오류: $($o.err)" }
  $n = @($o.rows).Count; if ($n -lt 1000) { throw "거래처가 $n 곳뿐이라 이상합니다(전체를 못 받은 듯)" }
  New-Item -ItemType Directory -Force (Split-Path $Out) | Out-Null
  $tmp = "$Out.tmp"; [IO.File]::WriteAllText($tmp, $json, (New-Object Text.UTF8Encoding $false)); Move-Item -LiteralPath $tmp $Out -Force
  "아마란스 거래처 $n 곳 → $Out"
} finally { Disconnect-Tab; Close-Tab $tab.id }
