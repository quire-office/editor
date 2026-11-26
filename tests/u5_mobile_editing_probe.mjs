/**
 * u5_mobile_editing_probe.mjs — 手機真觸控編輯鏈機械判據（用戶 11-26「手機無法編輯」）
 * 390x844@3x 仿真：真 touch 事件 → 四盞燈各自判紅：
 *  E1 touch 命中 canvas（版面吃掉可點區＝紅）
 *  E2 touch 後 focus 留在 inputSink（focus 被 Ribbon/HUD 偷＝紅）
 *  E3 touch 反查同段（caret_at 錯歸因＝紅）
 *  E4 IME 合成 commit 正確（合成期拼音混入文件＝紅；用戶實錄根因之一）
 * 用法：MOBILE_URL／CDP_URL／VW／VH／DPR env 覆寫。
 */
const CDP = process.env.CDP_URL || 'http://127.0.0.1:9222';
const URL_ = process.env.URL || 'http://127.0.0.1:8791/word.html';
const VW = Number(process.env.VW || 390), VH = Number(process.env.VH || 844), DPR = Number(process.env.DPR || 3);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const created = await (await fetch(`${CDP}/json/new?about:blank`, { method: 'PUT' })).json();
const sock = new WebSocket(created.webSocketDebuggerUrl);
await new Promise((res, rej) => { sock.onopen = res; sock.onerror = rej; });
let seq = 0; const pending = new Map(); const errors = [];
sock.onmessage = (e) => { const m = JSON.parse(e.data);
  if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); return; }
  if (m.method === 'Runtime.exceptionThrown') errors.push(String(m.params.exceptionDetails.exception?.description || '').slice(0, 140)); };
const send = (method, params = {}) => new Promise((res) => { const id = ++seq; pending.set(id, (m) => res(m.result || {})); sock.send(JSON.stringify({ id, method, params })); });
const ev = async (e) => { const r = await send('Runtime.evaluate', { expression: e, returnByValue: true, awaitPromise: true, timeout: 90000 }); if (r.exceptionDetails) return 'EX:' + String(r.exceptionDetails.exception?.description || '').slice(0, 200); return r.result?.value; };
await send('Page.enable'); await send('Runtime.enable');
await send('Network.enable'); await send('Network.setCacheDisabled', { cacheDisabled: true });
await send('Emulation.setDeviceMetricsOverride', { width: VW, height: VH, deviceScaleFactor: DPR, mobile: true });
await send('Emulation.setTouchEmulationEnabled', { enabled: true, maxTouchPoints: 5 });
await send('Page.navigate', { url: URL_ });
for (let i = 0; i < 200; i++) { if (await ev('document.readyState') === 'complete') break; await sleep(250); }
await ev('window.QWordShell.openDefaultSample()');
for (let i = 0; i < 200; i++) { if (await ev('!!window.QWordEngine.engine && window.QWordEngine.paragraphCount()>0')) break; await sleep(200); }
await sleep(800);
console.log(`== 手機仿真 v4 ${VW}x${VH}@${DPR}x  [${URL_}]`);
let red = 0;

await ev(`(function(){ window.__trace=[];
  ['touchstart','touchend','pointerdown','mousedown','focusin','focusout'].forEach(function(t){
    document.addEventListener(t, function(e){ window.__trace.push(t+':'+((e.target&&(e.target.id||e.target.className||e.target.tagName))||'?').toString().slice(0,26)); }, true); });
})()`);

// E0 fit-width 真算（U5-m：舊實作寫死 125%，手機上越按越大；用戶投訴版面根因）
await ev(`document.querySelector('[data-command="zoom:fit-width"]').click(); true`);
await sleep(600);
const fit = JSON.parse(await ev(`(function(){
  const c = document.querySelector('.qword-real-page canvas');
  const zw = document.querySelector('.qword-zoom-percent');
  return JSON.stringify({ cssW: Math.round(c.getBoundingClientRect().width), vw: innerWidth,
    zoom: zw ? zw.textContent : null });
})()`));
if (fit.cssW <= fit.vw) console.log(`✓ E0 fit-width 後頁寬 ${fit.cssW} ≤ 視窗 ${fit.vw}（${fit.zoom}）`);
else { console.log(`✗ E0 fit-width 後頁寬 ${fit.cssW} > 視窗 ${fit.vw}（${fit.zoom}）——仍塞不進`); red++; }

// 錨點：段 2（292 字元，行多）行中，必命中 canvas（v3 實錄 p2 y=282 hit canvas）
const G = JSON.parse(await ev(`(async function(){
  const A = window.QWordEngine; const p = 2;
  const t = Array.from(A.paragraphText(p)||''); const cp = 40;
  const w = document.querySelectorAll('.qword-real-page')[0];
  w.scrollIntoView({block:'start'});
  await new Promise(function(x){setTimeout(x,500);});
  const c = w.querySelector('canvas'); const rect = c.getBoundingClientRect();
  const dl = A.getPageDisplayList(0);
  const scaleCss = rect.width / (dl.width_pt || 595);
  const r = A.caretRect(p, cp);
  let x = Math.round(rect.left + (r.x_pt+1)*scaleCss); const y = Math.round(rect.top + (r.y_pt + r.height_pt - 4)*scaleCss); if (x > innerWidth-4) x = innerWidth-6;
  const el = document.elementFromPoint(x, y);
  return JSON.stringify({x:x, y:y, p:p, cp:cp, hit: el?String(el.className||el.tagName):null});
})()`));
console.log('tap 目標:', JSON.stringify(G));

// E1 touch 命中 canvas
if (G.hit && G.hit.includes('canvas')) console.log('✓ E1 touch 命中 canvas');
else { console.log('✗ E1 touch 命中=' + G.hit + '（版面吃掉了可點區）'); red++; }

await send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x: G.x, y: G.y, id: 1 }] });
await sleep(50);
await send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
await sleep(450);
const S = JSON.parse(await ev(`JSON.stringify({active: document.activeElement ? (document.activeElement.id||document.activeElement.tagName||document.activeElement.className) : null,
  caret: window.QWordEngine.caret, trace: window.__trace})`));
// E2 focus 歸屬
if (S.active === 'qword-input-sink') console.log('✓ E2 touch 後 focus=inputSink');
else { console.log('✗ E2 touch 後 focus=' + S.active + ' ← 鍵盤不會彈、打字必敗（手機無法編輯根因）'); red++; }
// E3 caret 反查
if (S.caret && S.caret.para === G.p) console.log('✓ E3 touch 反查同段 p=' + G.p);
else { console.log('✗ E3 caret=' + JSON.stringify(S.caret)); red++; }
console.log('  軌跡:', JSON.stringify(S.trace));

// E4 IME 合成路徑（真中文鍵盤管線：setComposition 拼音 → commit 中文）
if (S.active === 'qword-input-sink') {
  const before = await ev(`window.QWordEngine.paragraphText(${G.p})||''`);
  await send('Input.imeSetComposition', { text: 'zhong', selectionStart: 5, selectionEnd: 5 });
  await sleep(220);
  const mid = await ev(`JSON.stringify({para:(window.QWordEngine.paragraphText(${G.p})||''), sink: document.getElementById('qword-input-sink').value})`);
  await send('Input.insertText', { text: '中' });
  await sleep(320);
  const after = await ev(`window.QWordEngine.paragraphText(${G.p})||''`);
  const committed = String(after).includes('中');
  const pinyinLeak = String(after).includes('zhong');
  const grew = Array.from(String(after)).length - Array.from(String(before)).length;
  if (committed && !pinyinLeak && grew === 1) console.log('✓ E4 IME commit 正確（僅入「中」，+1 字元）');
  else { console.log(`✗ E4 IME 異常 committed=${committed} 拼音洩漏=${pinyinLeak} delta=${grew} mid=${mid}`); red++; }
}
if (errors.length) { console.log('✗ E5 JS 例外:', errors.slice(0, 2)); red++; }
console.log('\n判定：' + (red === 0 ? 'PASS' : 'FAIL（' + red + ' 項）'));
if (red) process.exitCode = 1;
try { await fetch(`${CDP}/json/close/${created.id}`); } catch {}
sock.close();
