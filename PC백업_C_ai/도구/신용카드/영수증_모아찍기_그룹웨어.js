/* 그룹웨어 지출결의 화면: '신용카드매출전표' 링크를 차례로 눌러 영수증을 모은 뒤 A4 한 장에 8장(2칸×4줄)으로 인쇄 화면을 띄움.
   사용: 영수증 목록이 보이는 화면에서 F12 → 콘솔(Console) → 이 내용을 붙여넣고 Enter */
(async () => {
  const LINK_TEXT = '신용카드매출전표';
  const MUST = ['승인번호', '거래일시'];            // 영수증 창에 꼭 있는 글자
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const log = (...a) => console.log('%c[영수증 모으기]', 'color:#0a6;font-weight:bold', ...a);
  const hasAll = el => el && MUST.every(w => (el.innerText || '').includes(w));

  // 1) 링크 찾기 (글자가 정확히 '신용카드매출전표'인 가장 안쪽 요소)
  const links = [...document.querySelectorAll('a,button,span,div,td,p,u,font')].filter(el =>
    (el.innerText || '').trim() === LINK_TEXT && ![...el.children].some(c => (c.innerText || '').trim() === LINK_TEXT));
  if (!links.length) { alert(`'${LINK_TEXT}' 링크를 못 찾았습니다. 영수증 목록이 보이는 화면에서 실행해 주세요.`); return; }
  log(`링크 ${links.length}개 찾음`);

  // 새 창(window.open)으로 뜨는 경우도 잡기
  const opened = []; const _open = window.open;
  window.open = function (...a) { const w = _open.apply(this, a); if (w) opened.push(w); return w; };

  const cards = []; const styleSrc = new Set();
  const closeLayer = async (layer) => {
    const btn = [...layer.querySelectorAll('button,a,span,div,i')].find(b => {
      const t = (b.innerText || '').trim(), c = (b.className && b.className.toString()) || '', l = b.getAttribute('aria-label') || b.getAttribute('title') || '';
      return /^(닫기|확인|×|X|x|취소)$/.test(t) || /close|btn_x|btnClose|popClose/i.test(c) || /닫기|close/i.test(l);
    });
    if (btn) btn.click(); else document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true }));
    for (let i = 0; i < 20 && document.contains(layer) && layer.offsetParent !== null; i++) await sleep(100);
  };
  const cleanClone = (node) => {
    const c = node.cloneNode(true);
    c.querySelectorAll('button,script').forEach(b => b.remove());
    [c, ...c.querySelectorAll('*')].forEach(e => { if (e.style) { e.style.position = 'static'; e.style.transform = 'none'; e.style.maxHeight = 'none'; e.style.overflow = 'visible'; } });
    return c;
  };

  for (let i = 0; i < links.length; i++) {
    const before = new Set(document.querySelectorAll('body *'));
    const added = []; const mo = new MutationObserver(ms => ms.forEach(m => m.addedNodes.forEach(n => n.nodeType === 1 && added.push(n))));
    mo.observe(document.body, { childList: true, subtree: true });
    const nWin = opened.length;
    links[i].scrollIntoView({ block: 'center' }); links[i].click();
    let got = null, layer = null, win = null;
    for (let t = 0; t < 60 && !got; t++) {           // 최대 6초 기다림
      await sleep(100);
      if (opened.length > nWin) {                       // 새 창으로 뜬 경우
        win = opened[opened.length - 1];
        try { if (win.document && hasAll(win.document.body)) { got = win.document.body; win.document.querySelectorAll('link[rel=stylesheet],style').forEach(s => styleSrc.add(s.outerHTML)); } } catch (e) { }
        continue;
      }
      // 같은 화면 안 레이어로 뜬 경우: 새로 생긴(또는 다시 보이는) 요소 중 영수증 글자가 있는 가장 바깥
      const cand = added.filter(n => document.contains(n) && hasAll(n));
      if (cand.length) { layer = cand.reduce((a, b) => (a.contains(b) ? a : b.contains(a) ? b : a)); got = layer; break; }
      const vis = [...document.querySelectorAll('[role=dialog],.modal,.popup,.layer,[class*=Popup],[class*=Dialog],[class*=Modal],[class*=layer]')]
        .filter(e => e.offsetParent !== null && hasAll(e));
      if (vis.length) { layer = vis[vis.length - 1]; got = layer; }
    }
    mo.disconnect();
    if (!got) { log(`${i + 1}번째: 영수증 창을 못 찾음 (건너뜀)`); continue; }
    // 영수증 본문만: '사업자등록번호'~'거래일시'를 다 가진 가장 작은 상자 → 가게 이름·금액까지 포함되게 한 단계씩 넓힘
    let body = got;
    const inner = [...got.querySelectorAll('*')].filter(hasAll);
    if (inner.length) { body = inner[inner.length - 1]; while (body.parentElement && body !== got && !/\d[\d,]*\s*원/.test(body.innerText)) body = body.parentElement; }
    cards.push(cleanClone(body).outerHTML);
    log(`${i + 1}/${links.length} 모음`);
    if (win) { try { win.close(); } catch (e) { } } else if (layer) await closeLayer(layer);
    await sleep(250);
  }
  window.open = _open;
  if (!cards.length) { alert('영수증을 하나도 못 모았습니다. 콘솔에 나온 내용을 알려 주세요.'); return; }

  // 2) A4 한 장에 8장 인쇄 화면
  styleSrc.forEach(h => document.head.insertAdjacentHTML('beforeend', h));
  document.getElementById('rcpt8') && document.getElementById('rcpt8').remove();
  const css = `
    #rcpt8{position:fixed;inset:0;z-index:2147483647;background:#888;overflow:auto;padding:12px}
    #rcpt8 .bar{position:sticky;top:0;background:#fff;padding:8px 12px;margin:0 auto 10px;width:210mm;box-sizing:border-box;font:14px sans-serif;display:flex;gap:10px;align-items:center}
    #rcpt8 .bar button{font-size:14px;padding:6px 14px;cursor:pointer}
    #rcpt8 .pg{width:210mm;height:297mm;background:#fff;margin:0 auto 10px;padding:6mm;box-sizing:border-box;display:grid;grid-template-columns:1fr 1fr;grid-template-rows:repeat(4,1fr);gap:0}
    #rcpt8 .cell{border:1px dashed #bbb;overflow:hidden;display:flex;align-items:center;justify-content:center}
    #rcpt8 .in{display:inline-block;background:#fff}
    @media print{
      @page{size:A4;margin:0}
      body>*:not(#rcpt8){display:none!important}
      #rcpt8{position:static;background:none;padding:0;overflow:visible}
      #rcpt8 .bar{display:none}
      #rcpt8 .pg{margin:0;page-break-after:always;break-after:page}
    }`;
  const wrap = document.createElement('div'); wrap.id = 'rcpt8';
  let html = `<style>${css}</style><div class="bar">영수증 ${cards.length}장 → A4 ${Math.ceil(cards.length / 8)}장
    <button id="rcpt8print">인쇄</button><button id="rcpt8close">닫기</button>
    <span style="color:#666">인쇄 창에서 '배율: 기본값', '여백: 없음'</span></div>`;
  for (let p = 0; p < cards.length; p += 8) html += `<div class="pg">${cards.slice(p, p + 8).map(c => `<div class="cell"><div class="in">${c}</div></div>`).join('')}${'<div class="cell"></div>'.repeat(Math.max(0, 8 - cards.slice(p, p + 8).length))}</div>`;
  wrap.innerHTML = html; document.body.appendChild(wrap);
  await sleep(300);
  wrap.querySelectorAll('.cell').forEach(cell => {      // 칸 크기에 맞게 줄이기
    const inn = cell.querySelector('.in'); if (!inn) return;
    const z = Math.min((cell.clientWidth - 8) / inn.scrollWidth, (cell.clientHeight - 8) / inn.scrollHeight, 1.5);
    inn.style.zoom = z;
  });
  document.getElementById('rcpt8print').onclick = () => window.print();
  document.getElementById('rcpt8close').onclick = () => wrap.remove();
  log(`완료: ${cards.length}장`);
})();
