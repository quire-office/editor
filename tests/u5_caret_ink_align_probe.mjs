/**
 * u5_caret_ink_align_probe.mjs — caret vs 文字墨跡同框對齊機械判據（用戶 11-26「比文字小且偏上」）
 * 同框絕對座標量測 DOM caret 上下緣 vs canvas 像素墨跡上下緣（clean 行窗不跨行）。
 * 判讀：dTop（墨跡頂−caret 頂）應 ≈0（±2px）且 caret 完整覆蓋墨跡（dBot≥0）。
 * 桌面態與手機 fit-width 態（FIT=1）雙態可測；跨字級自動適用（墨跡真值尺同源）。
 */
const CDP = process.env.CDP_URL || 'http://127.0.0.1:9222';
const URL_ = process.env.URL || 'http://127.0.0.1:8791/word.html';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const VW = Number(process.env.VW || 0), VH = Number(process.env.VH || 0), DPR = Number(process.env.DPR || 1);
const created = await (await fetch(`${CDP}/json/new?about:blank`, { method: 'PUT' })).json();
const sock = new WebSocket(created.webSocketDebuggerUrl);
await new Promise((res, rej) => { sock.onopen = res; sock.onerror = rej; });
let seq = 0; const pending = new Map();
sock.onmessage = (e) => { const m = JSON.parse(e.data); if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } };
const send = (method, params = {}) => new Promise((res) => { const id = ++seq; pending.set(id, (m) => res(m.result || {})); sock.send(JSON.stringify({ id, method, params })); });
const ev = async (e) => { const r = await send('Runtime.evaluate', { expression: e, returnByValue: true, awaitPromise: true, timeout: 90000 }); if (r.exceptionDetails) return 'EX:' + String(r.exceptionDetails.exception?.description || '').slice(0, 200); return r.result?.value; };
await send('Page.enable'); await send('Runtime.enable');
await send('Network.enable'); await send('Network.setCacheDisabled', { cacheDisabled: true });
if (VW) {
  await send('Emulation.setDeviceMetricsOverride', { width: VW, height: VH, deviceScaleFactor: DPR, mobile: true });
  await send('Emulation.setTouchEmulationEnabled', { enabled: true, maxTouchPoints: 5 });
}
await send('Page.navigate', { url: URL_ });
for (let i = 0; i < 200; i++) { if (await ev('document.readyState') === 'complete') break; await sleep(250); }
await ev('window.QWordShell.openDefaultSample()');
for (let i = 0; i < 200; i++) { if (await ev('!!window.QWordEngine.engine && window.QWordEngine.paragraphCount()>0')) break; await sleep(200); }
await sleep(800);
if (process.env.FIT) { await ev("document.querySelector('[data-command=\"zoom:fit-width\"]').click(); true"); await sleep(700); }
const out = await ev(`(function(){
  const A = window.QWordEngine;
  const rows = [];
  for (const p of [1, 2, 3]) {
    const t = Array.from(A.paragraphText(p)||'');
    if (t.length < 30) continue;
    const cp = 20;
    A.caret = { para: p, cp: cp }; A.selection = null;
    window.QWordShell.render();
    const r = A.caretRect(p, cp);
    const w = document.querySelectorAll('.qword-real-page')[r.page];
    const c = w.querySelector('canvas');
    const dl = A.getPageDisplayList(r.page);
    const scale = c.width / (dl.width_pt || 595);      // backing px / pt
    const cssPerBack = c.getBoundingClientRect().width / c.width; // CSS/DOM px per backing px
    const el = w.querySelector('.qword-caret');
    // DOM caret 絕對 DOM px（相對 canvas 框）
    const cRect = c.getBoundingClientRect(), eRect = el.getBoundingClientRect();
    const caretTopDom = Math.round((eRect.top - cRect.top)*10)/10, caretBotDom = Math.round((eRect.bottom - cRect.top)*10)/10;
    // canvas 像素墨跡（絕對 backing px）：caret 左側 60pt 帶、caret 上下 ±14pt 窗
    const ctx = c.getContext('2d');
    const X0 = Math.max(0, Math.round((r.x_pt-60)*scale)), X1 = Math.round(r.x_pt*scale);
    const Y0 = Math.max(0, Math.round(r.y_pt*scale) - 3), Y1 = Math.round((r.y_pt + r.height_pt + 5)*scale);
    const img = ctx.getImageData(X0, Y0, X1-X0, Y1-Y0).data;
    const rowsInk = [];
    for (let y=0;y<(Y1-Y0);y++){ let n=0;
      for (let x=0;x<(X1-X0)*4;x+=4){ const i=y*(X1-X0)*4+x; if (img[i]<110&&img[i+1]<110&&img[i+2]<110&&img[i+3]>90) n++; }
      if (n>=3) rowsInk.push(Y0+y); }
    const inkTopB = rowsInk.length?rowsInk[0]:null, inkBotB = rowsInk.length?rowsInk[rowsInk.length-1]:null;
    // backing→DOM：墨跡換算到 DOM px 系（與 caretTopDom 同框）
    const inkTopDom = inkTopB!==null?Math.round((inkTopB)*cssPerBack*10)/10:null;
    const inkBotDom = inkBotB!==null?Math.round((inkBotB)*cssPerBack*10)/10:null;
    rows.push({ p, size: r.height_pt, caretTopDom, caretBotDom, caretH: Math.round((caretBotDom-caretTopDom)*10)/10,
                inkTopDom, inkBotDom, inkH: inkTopDom!==null?Math.round((inkBotDom-inkTopDom)*10)/10:null,
                dTop: inkTopDom!==null?Math.round((inkTopDom-caretTopDom)*10)/10:null,
                dBot: inkBotDom!==null?Math.round((caretBotDom-inkBotDom)*10)/10:null });
  }
  return JSON.stringify({ vw: innerWidth, dpr: devicePixelRatio, canvasW: document.querySelector('.qword-real-page canvas').width, cssW: Math.round(document.querySelector('.qword-real-page canvas').getBoundingClientRect().width), rows: rows });
})()`);
const J = JSON.parse(out);
console.log(`[vw=${J.vw} dpr=${J.dpr} canvasBacking=${J.canvasW} cssW=${J.cssW}]`);
console.log('p  size caretTop-Bot      inkTop-Bot         caretH inkH  dTop(ink-caret) dBot(caret-ink)');
for (const r of J.rows) console.log(`${r.p} ${r.size} [${r.caretTopDom},${r.caretBotDom}] [${r.inkTopDom},${r.inkBotDom}] ${r.caretH} ${r.inkH}  ${r.dTop} ${r.dBot}`);
// 判定：頂對墨跡（|dTop|≤2px）＋ caret 完整蓋住墨跡（dBot≥−1，留 1px 取樣容差）
let bad = 0;
for (const r of J.rows) {
  if (r.dTop === null) continue;
  if (Math.abs(r.dTop) > 2 || r.dBot < -1) bad++;
}
console.log(`\n判定：${bad === 0 ? 'PASS（caret 錨定墨跡框）' : 'FAIL（' + bad + ' 段偏位）'}`);
if (bad) process.exitCode = 1;
try { await fetch(`${CDP}/json/close/${created.id}`); } catch {}
sock.close();
