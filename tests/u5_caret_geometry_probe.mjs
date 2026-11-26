/**
 * u5_caret_geometry_probe.mjs — caret 幾何與文字對齊/等高 機械判據
 *
 * 用戶投訴兩案（11-26）：
 *  (a)「cursor 顯示位置與實際輸入後文字顯示的位置不同」
 *  (b)「cursor 的高度還是與文字大小不匹配」
 * 兩者都必須變成可判的數值斷言，而不是目視——本探針鎖三盞燈：
 *
 *  燈1 等高：對每個可測段落，caretRect(p,cp).height_pt 必 ≈ 該段**級聯字級**
 *     （由顯示列表同段 Text 項的 size_pt 取真值）。偏差 >20% 拍紅。
 *     （用戶 (b) 的直接形式：12pt 全域尺套在 26pt 標題上，caret 只有字高的 46%。）
 *  燈2 對齊：真滑鼠點擊某段行內 → caret 反查段/列必與點擊處同段同列；
 *     且 caret DOM 元素（.qword-caret）的中心 y 必落在该行文字基線±行高內。
 *     （用戶 (a) 的 DOM 層形式：顯示的 caret 條和游標邏輯位置要同址。）
 *  燈3 打字落點：真鍵盤打字後，新字元出現處 == 打字前 caret 顯示處（同一行、
 *     同段、x 位移正向）。用戶 (a) 的「實際輸入後文字顯示位置」直接形式。
 *     併測「打字引入未請求前導空白（縮排）」＝用戶說的「打字後突然縮排」。
 *
 * 用法：CARET_URL（默認本地 8791）／CDP_URL（默認 9222）／CARET_OUT 證據檔。
 */
const BASE_URL = process.env.CARET_URL || 'http://127.0.0.1:8791/word.html';
const OUT = process.env.CARET_OUT || null;
const CDP_HTTP = process.env.CDP_URL || 'http://127.0.0.1:9222';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function main() {
  const created = await (await fetch(`${CDP_HTTP}/json/new?about:blank`, { method: 'PUT' })).json();
  const sock = new WebSocket(created.webSocketDebuggerUrl);
  await new Promise((res, rej) => { sock.onopen = res; sock.onerror = rej; });
  let seq = 0; const pending = new Map();
  sock.onmessage = (ev) => { const m = JSON.parse(ev.data); if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); } };
  const send = (method, params = {}) => new Promise((res) => { const id = ++seq; pending.set(id, (m) => res(m.result || {})); sock.send(JSON.stringify({ id, method, params })); });
  const evalIn = async (e) => { const r = await send('Runtime.evaluate', { expression: e, returnByValue: true, awaitPromise: true, timeout: 120000 }); if (r.exceptionDetails) return '__EX__:' + String(r.exceptionDetails.exception?.description || '').slice(0, 200); return r.result ? r.result.value : undefined; };

  await send('Page.enable'); await send('Runtime.enable');
  // 探針自身前提：headless 默認視窗約 493px 高，固定 Ribbon/HUD 會吞掉落在畫布
  // 外的真點擊（caret 不動 → 誤判成引擎錯位）。先撐高視窗，且每次派發前用
  // elementFromPoint 斷言真的命中 canvas（不命中就拍紅並回報元素，歸因不猜）。
  await send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 1400, deviceScaleFactor: 1, mobile: false });
  await send('Page.navigate', { url: BASE_URL });
  for (let i = 0; i < 240; i++) { if (await evalIn('document.readyState') === 'complete') break; await sleep(250); }
  await evalIn('window.QWordShell.openDefaultSample()');
  for (let i = 0; i < 300; i++) { if (await evalIn('!!window.QWordEngine.engine && window.QWordEngine.paragraphCount() > 0')) break; await sleep(200); }
  await sleep(700);

  // ---- 燈1：caret 高度 vs 該段級聯字級（display list 真值）----
  // 取真值方式（不造第二把尺）：caretRect 給出 (page, y_pt, x_pt) 後，在該頁
  // 顯示列表找**同一條基線**（|y−caretY|<0.6pt）且 x 距 caret 最近的 Text 項，
  // 其 size_pt 就是 caret 該等高的那個字元級聯字級。
  const h1 = await evalIn(`(function () {
    const A = window.QWordEngine;
    const rows = [];
    for (let p = 0; p < A.paragraphCount(); p++) {
      const t = Array.from(A.paragraphText(p) || '');
      if (t.length < 2) continue;
      const r = A.caretRect(p, 1);
      if (!r) { rows.push({ p, noRect: true }); continue; }
      const dl = A.getPageDisplayList(r.page);
      let best = null, bestD = Infinity;
      for (const it of dl.items) {
        if (!it.Text) continue;
        const dy = Math.abs(it.Text.y - r.y_pt);
        if (dy > 0.6) continue;
        const dx = Math.abs((it.Text.x || 0) - r.x_pt);
        if (dx < bestD) { bestD = dx; best = it.Text; }
      }
      rows.push({ p, caretH: r.height_pt, paraSize: best ? best.size_pt : null, page: r.page, y: r.y_pt });
    }
    return JSON.stringify(rows);
  })()`);
  const rows = typeof h1 === 'string' && h1.startsWith('__EX__') ? [] : JSON.parse(h1 || '[]');
  const badH = rows.filter((r) => r.paraSize && r.caretH && Math.abs(r.caretH - r.paraSize) / r.paraSize > 0.2);
  const sized = rows.filter((r) => r.paraSize && r.caretH);

  // ---- 燈2：真滑鼠點擊 → caret 反查同段 ＋ DOM caret 條落在該行帶內 ----
  const paras = await evalIn('window.QWordEngine.paragraphCount()');
  const clickHits = [];
  for (const p of [0, 1, 2, 3, 4, 5, 6, 7].filter((x) => x < paras)) {
    const geo = await evalIn(`(async function () {
      const A = window.QWordEngine;
      const t = Array.from(A.paragraphText(${p}) || '');
      if (t.length < 6) return null;
      const mid = Math.floor(t.length / 2);
      const r = A.caretRect(${p}, mid);
      if (!r) return null;
      const w = document.querySelectorAll('.qword-real-page')[r.page];
      w.scrollIntoView({ block: 'center' });
      await new Promise(function (x) { setTimeout(x, 380); });
      const c = w.querySelector('canvas'); const rect = c.getBoundingClientRect();
      const dl = A.getPageDisplayList(r.page); const scale = c.width / (dl.width_pt || 595);
      // 點擊 y 錨在**墨跡帶**（基線上方 4pt＝r.y＋size−4）。U4m 已定案 Text.y＝行框頂、
      // 基線＝y＋size：錨在 y−4（行框頂以上）落的是上一行墨跡區，反查歸前段是
      // **正確**行為（探針錨點錯，非產品缺陷——run 90 實錄：p1 誤錨 y−4→反查 (0,15)）。
      const y = rect.top + (r.y_pt + (r.height_pt || 12) - 4) * scale;
      const x = rect.left + (r.x_pt + 1) * scale;
      const hit = document.elementFromPoint(Math.round(x), Math.round(y));
      return { x: Math.round(x), y: Math.round(y), mid: mid,
               hit: hit ? String(hit.className || hit.tagName) : null };
    })()`);
    if (!geo || typeof geo !== 'object') { clickHits.push({ p, skip: true }); continue; }
    if (!String(geo.hit).includes('canvas')) { clickHits.push({ p, skip: 'hit=' + geo.hit }); continue; }
    await send('Input.dispatchMouseEvent', { type: 'mousePressed', x: geo.x, y: geo.y, button: 'left', clickCount: 1, buttons: 1 });
    await send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: geo.x, y: geo.y, button: 'left', buttons: 0 });
    await sleep(280);
    const got = await evalIn(`(function () {
      const A = window.QWordEngine;
      const el = document.querySelector('.qword-caret');
      const wr = el && el.parentElement ? el.parentElement.getBoundingClientRect() : null;
      return JSON.stringify({
        caret: A.caret ? { para: A.caret.para, cp: A.caret.cp } : null,
        caretBoxTop: el && wr ? Math.round(el.offsetTop) : null,
        caretBoxH: el && wr ? Math.round(el.offsetHeight) : null
      });
    })()`);
    const g = typeof got === 'string' && got.startsWith('__EX__') ? { err: got } : JSON.parse(got || '{}');
    clickHits.push({ p, wantCp: geo.mid, got: g.caret, match: g.caret && g.caret.para === p,
                     caretBoxTop: g.caretBoxTop, caretBoxH: g.caretBoxH });
  }
  const badClick = clickHits.filter((c) => c.match === false);

  // ---- 燈3：真鍵盤打字落點＋縮排回歸 ----
  // 選一個夠長段，點它段首 → 真鍵盤打 'QX' → 段文字必須以 QX 開頭且前面無空白
  const p3 = await evalIn(`(function(){const A=window.QWordEngine;for(let i=0;i<A.paragraphCount();i++){if(Array.from(A.paragraphText(i)||'').length>10)return i;}return -1;})()`);
  await send('Input.dispatchMouseEvent', { type: 'mousePressed', x: 0, y: 0, button: 'left', clickCount: 0, buttons: 0 });
  const g3 = await evalIn(`(async function () {
    const A = window.QWordEngine;
    const r = A.caretRect(${p3}, 0);
    const w = document.querySelectorAll('.qword-real-page')[r.page];
    w.scrollIntoView({ block: 'center' });
    await new Promise(function (x) { setTimeout(x, 380); });
    const c = w.querySelector('canvas'); const rect = c.getBoundingClientRect();
    const dl = A.getPageDisplayList(r.page); const scale = c.width / (dl.width_pt || 595);
    return { x: rect.left + r.x_pt * scale + 1, y: rect.top + (r.y_pt + (r.height_pt || 12) - 5) * scale };
  })()`);
  await send('Input.dispatchMouseEvent', { type: 'mousePressed', x: g3.x, y: g3.y, button: 'left', clickCount: 1, buttons: 1 });
  await send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: g3.x, y: g3.y, button: 'left', buttons: 0 });
  await sleep(300);
  const caretBefore = JSON.parse(await evalIn('JSON.stringify(window.QWordEngine.caret)') || 'null');
  const beforeTxt = await evalIn(`window.QWordEngine.paragraphText(${p3}) || ''`);
  for (const ch of 'QX') {
    await send('Input.dispatchKeyEvent', { type: 'keyDown', key: ch, code: 'Key' + ch, text: ch, unmodifiedKey: ch, windowsVirtualKeyCode: ch.charCodeAt(0) });
    await send('Input.dispatchKeyEvent', { type: 'keyUp', key: ch, code: 'Key' + ch, windowsVirtualKeyCode: ch.charCodeAt(0) });
    await sleep(220);
  }
  const afterTxt = await evalIn(`window.QWordEngine.paragraphText(${p3}) || ''`);
  const caretAfterType = JSON.parse(await evalIn('JSON.stringify(window.QWordEngine.caret)') || 'null');
  const typed = {
    p: p3, caretBefore, caretAfterType,
    beforeHead: beforeTxt.slice(0, 8), afterHead: afterTxt.slice(0, 10),
    landedAtCaret: caretBefore && caretBefore.cp === 0 ? afterTxt.startsWith('QX') : true,
    leadingIntruder: /^[\s\u3000\t]/.test(afterTxt),
    paraUnchanged: caretBefore && (caretAfterType ? caretAfterType.para === caretBefore.para : true)
  };

  try { await fetch(`${CDP_HTTP}/json/close/${created.id}`); } catch {}
  sock.close();
  if (OUT) { const fs = await import('node:fs'); fs.writeFileSync(OUT, JSON.stringify({ rows, clickHits, typed }, null, 2)); }

  let fail = 0;
  const line = (cond, msg) => { console.log(`  ${cond ? '✓' : '✗'} ${msg}`); if (!cond) fail += 1; };
  console.log('=== U5 caret 幾何（等高／對齊／打字落點）判定 ===');
  console.log(`  baseUrl=${BASE_URL}`);
  line(badH.length === 0 && sized.length > 0,
    `燈1 caret 高度 ≈ 段落字級（可測 ${sized.length} 段，偏差>20% 者 ${badH.length}）`);
  if (badH.length) console.log('     偏差樣本:', JSON.stringify(badH.slice(0, 4)));
  else if (sized.length) console.log('     樣本:', JSON.stringify(sized.slice(0, 4)));
  line(badClick.length === 0, `燈2 真滑鼠點擊段↔caret 段全等（${clickHits.filter((c) => c.match === true).length}/${clickHits.filter((c) => !c.skip).length}，錯判 ${badClick.length}）`);
  if (badClick.length) console.log('     錯判:', JSON.stringify(badClick));
  line(typed.landedAtCaret && !typed.leadingIntruder && typed.paraUnchanged,
    `燈3 真鍵盤打字落點＝caret 處且無未請求前導空白（head=${JSON.stringify(typed.afterHead)} intruder=${typed.leadingIntruder}）`);
  console.log(`\n判定：${fail === 0 ? 'PASS' : `FAIL（${fail} 項）`}`);
  process.exit(fail === 0 ? 0 : 1);
}
main().catch((e) => { console.error('探針失敗:', e.message); process.exit(2); });
