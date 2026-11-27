/**
 * u5_n_font_ruler_probe.mjs — U5-n：渲染/排版同尺（caret 飄移案）＋ DPR 高清（模糊案）
 *
 * 用戶實錄（11-26）：①「cursor 位置會飄到超級前面」（手機打字後 caret 與墨跡分離
 * 累積）②「開啟的檔案編輯畫面有點模糊」。兩根因：
 *  R1 雙尺：渲染器用瀏覽器字型（手機無 Noto Sans TC → 落回系統字型）畫字，
 *     caret x 卻用引擎內嵌字型 advance → 打字越長飄越遠。修復＝引擎出口
 *     embeddedFontBytes/Family，殼端註冊 FontFace，渲染器與 caret 度量堆疊
 *     第一順位同家族（「Quire Noto Sans TC」）。
 *  R2 模糊：canvas backing 固定 794px，手機 DPR=3 被拉伸 3 倍顯示。修復＝
 *     backing ×dpr、CSS 尺寸守 794 版面、ctx.setTransform(dpr)。
 *
 * 三燈（全部 DPR=3 手機仿真態）：
 *  E1 字型註冊：document.fonts.check(12px "Quire Noto Sans TC") 為真，且同串
 *     CJK+拉丁混排的 measureText 寬度 == 引擎 caret advance（±1%，同尺證明）。
 *  E2 caret 緊貼：真 touch 點擊段中 → 真 IME 打 15 字混排 → caret DOM 左緣對
 *     「左側最近墨跡右緣」gap ≤ 4 CSS px（舊 bug 態＝飄 100px+）。
 *  E3 高清：canvas.width == round(offsetWidth × dpr)（backing 跟上 DPR），且
 *     版面 CSS 寬仍 = 794（佈局未跑版）。
 */
const CDP = process.env.CDP_URL || 'http://127.0.0.1:9222';
const URL_ = process.env.URL || 'http://127.0.0.1:8791/word.html';
const VW = Number(process.env.VW || 390), VH = Number(process.env.VH || 844), DPR = Number(process.env.DPR || 3);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const fails = [];
const ok = (cond, msg) => { console.log(`${cond ? '✓' : '✗'} ${msg}`); if (!cond) fails.push(msg); };

const created = await (await fetch(`${CDP}/json/new?about:blank`, { method: 'PUT' })).json();
const sock = new WebSocket(created.webSocketDebuggerUrl);
await new Promise((res, rej) => { sock.onopen = res; sock.onerror = rej; });
let seq = 0; const pending = new Map();
sock.onmessage = (e) => { const m = JSON.parse(e.data); if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } };
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

// ---- E1 同尺：FontFace 註冊＋瀏覽器度量 == 引擎 advance ----
const E1 = JSON.parse(await ev(`(function(){
  const fam = 'Quire Noto Sans TC';
  const registered = document.fonts.check('12px "' + fam + '"');
  const A = window.QWordEngine;
  // 引擎尺：段 2 前 45 字元（拉丁混排）的 caret advance 總和（pt）
  const t = Array.from(A.paragraphText(2) || '');
  const n = 45, s = t.slice(0, n).join('');
  const adv = A.caretRect(2, n).x_pt - A.caretRect(2, 0).x_pt;
  // 瀏覽器尺：measureText 不受 ctx transform 影響、回傳 CSS px；12pt 要按
  // 96/72 換 px（16px），結果再乘 0.75 回 pt——與引擎 pt 同单位的對比。
  const c = document.querySelector('.qword-real-page canvas');
  const ctx = c.getContext('2d'); ctx.font = '16px "' + fam + '"';   // 12pt = 16px
  const browserW = ctx.measureText(s).width * 0.75;
  return JSON.stringify({ registered, adv: Math.round(adv*100)/100, browserW: Math.round(browserW*100)/100 });
})()`));
ok(E1.registered, `E1a 內嵌字型已註冊（document.fonts.check）`);
const ratio = E1.browserW / E1.adv;
ok(Math.abs(ratio - 1) <= 0.01, `E1b 同尺：瀏覽器/引擎寬比 ${ratio.toFixed(4)}（引擎 ${E1.adv}pt vs 瀏覽器 ${E1.browserW}pt，容差 1%）`);

// ---- E3 高清：backing = CSS×DPR；版面 CSS 寬不跑版 ----
const E3 = JSON.parse(await ev(`(function(){
  const c = document.querySelector('.qword-real-page canvas');
  return JSON.stringify({ backing: c.width, offsetW: c.offsetWidth, dpr: devicePixelRatio });
})()`));
ok(Math.abs(E3.backing - Math.round(E3.offsetW * DPR)) <= 2, `E3a backing ${E3.backing} == CSS ${E3.offsetW} × DPR ${DPR}`);
ok(Math.abs(E3.offsetW - 794) <= 2, `E3b 版面 CSS 寬 ${E3.offsetW}（==794，佈局未跑版）`);

// ---- E2 caret 緊貼：真 touch 點段 → 真 IME 打字 → caret vs 墨跡 ----
const P = 2;
const G = JSON.parse(await ev(`(async function(){
  const A = window.QWordEngine; const p=${P};
  const t=Array.from(A.paragraphText(p)||''); const cp=Math.floor(t.length/2);
  A.caret={para:p,cp:cp}; A.selection=null; window.QWordShell.render();
  const r=A.caretRect(p,cp);
  const w=document.querySelectorAll('.qword-real-page')[r.page];
  w.scrollIntoView({block:'center'}); await new Promise(function(x){setTimeout(x,420);});
  const c=w.querySelector('canvas'); const rect=c.getBoundingClientRect();
  const dl=A.getPageDisplayList(r.page); const scaleCss=rect.width/(dl.width_pt||595);
  let x=Math.round(rect.left+(r.x_pt+1)*scaleCss); const y=Math.round(rect.top+(r.y_pt+r.height_pt-4)*scaleCss);
  if (x>innerWidth-6) x=innerWidth-6;
  return JSON.stringify({x:x,y:y});
})()`));
await send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x: G.x, y: G.y, id: 1 }] });
await sleep(50);
await send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
await sleep(350);
await send('Input.imeSetComposition', { text: '因為我', selectionStart: 3, selectionEnd: 3 });
await sleep(120);
await send('Input.insertText', { text: '因為我' });
await sleep(220);
for (const ch of 'ytfgytdigudyff') { await send('Input.insertText', { text: ch }); await sleep(80); }
await sleep(500);

const E2 = JSON.parse(await ev(`(function(){
  const A = window.QWordEngine;
  const cA=A.caret;
  const r=A.caretRect(cA.para,cA.cp);
  const w=document.querySelectorAll('.qword-real-page')[r.page];
  const c=w.querySelector('canvas'); const dl=A.getPageDisplayList(r.page);
  const scale=c.width/(dl.width_pt||595);
  const cssPerBack=(c.offsetWidth||c.width)/c.width;
  const el=w.querySelector('.qword-caret');
  const ctx=c.getContext('2d');
  const Y0=Math.max(0,Math.round(r.y_pt*scale)-3), Y1=Math.round((r.y_pt+r.height_pt+5)*scale);
  const X0=Math.max(0,Math.round((r.x_pt-80)*scale)), X1=Math.min(c.width-1,Math.round(r.x_pt*scale));
  const img=ctx.getImageData(X0,Y0,X1-X0,Y1-Y0).data;
  let inkR=-1;
  for(let x=(X1-X0)-1;x>=0;x--){ let n=0;
    for(let y=0;y<(Y1-Y0);y++){ const i=(y*(X1-X0)+x)*4; if(img[i]<110&&img[i+1]<110&&img[i+2]<110&&img[i+3]>90){n++; if(n>2)break;} }
    if(n>2){ inkR=X0+x; break; } }
  return JSON.stringify({ caretCp:cA.cp, gapPx: inkR>=0?Math.round((el.offsetLeft-inkR*cssPerBack)*10)/10:null });
})()`));
ok(E2.gapPx !== null && Math.abs(E2.gapPx) <= 4, `E2 打字 18 字後 caret 對墨跡 gap=${E2.gapPx}px（≤4；舊 bug 態飄 100px+）`);

console.log(`\n判定：${fails.length === 0 ? 'PASS' : 'FAIL'}（${fails.length} 項紅）`);
try { await fetch(`${CDP}/json/close/${created.id}`); } catch {}
sock.close();
process.exit(fails.length ? 1 : 0);
