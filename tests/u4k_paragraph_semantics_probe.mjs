// U4k 段落級編輯閉環探針（跨段 Enter/Backspace＋選區鍵盤操作）
// 鑑別力紀律：全部走 CDP 真鍵盤／真滑鼠事件路徑（u4h 教訓），驗收前做
// BUGGY 態驗證（暫時摘除修復→對應燈必紅）。
// 用法：node tests/u4k_paragraph_semantics_probe.mjs（自起 headless Chrome）
//   URL（預帶 http://127.0.0.1:8791/word.html）、U4K_CDP_PORT（預帶 9228）。
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';

const BASE = (process.env.U4K_URL || 'http://127.0.0.1:8791/word.html').replace(/\/$/, '');
const CHROME = process.env.CHROME_BIN || '/usr/bin/google-chrome';
const CDP_PORT = process.env.U4K_CDP_PORT || '9228';
const CDP_HTTP = `http://127.0.0.1:${CDP_PORT}`;
const OUT = process.env.U4K_OUT || path.join('evidence', 'word-ui-u4k-paragraph-evidence.json');

class CDP {
  constructor(ws) { this.ws = ws; this.id = 0; this.pending = new Map();
    ws.addEventListener('message', (m) => { const d = JSON.parse(m.data);
      if (d.id && this.pending.has(d.id)) { this.pending.get(d.id)(d); this.pending.delete(d.id); } }); }
  send(method, params = {}) { return new Promise((resolve, reject) => { const i = ++this.id;
    this.pending.set(i, resolve); this.ws.send(JSON.stringify({ id: i, method, params }));
    setTimeout(() => { if (this.pending.has(i)) { this.pending.delete(i); reject(new Error(`CDP 逾時: ${method}`)); } }, 20000); }); }
}
const results = [];
const fail = (n, d) => { results.push({ n, ok: false, d }); console.log(`  ✗ ${n}｜${d}`); };
const pass = (n, d) => { results.push({ n, ok: true, d }); console.log(`  ✓ ${n}｜${d}`); };

const profileDir = fs.mkdtempSync(path.join(os.tmpdir(), 'u4k-probe-'));
const chromeChild = spawn(CHROME, ['--headless=new', `--user-data-dir=${profileDir}`, `--remote-debugging-port=${CDP_PORT}`,
  '--no-first-run', '--no-default-browser-check', '--disable-gpu', 'about:blank'], { stdio: 'ignore' });
let cdp = null; let ws = null;
for (let i = 0; i < 40 && !cdp; i++) {
  await new Promise((r) => setTimeout(r, 400));
  try {
    const list = await (await fetch(`${CDP_HTTP}/json/list`)).json();
    const t = list.find((x) => x.type === 'page');
    if (!t) continue;
    ws = new WebSocket(t.webSocketDebuggerUrl);
    cdp = await new Promise((resolve, reject) => {
      ws.addEventListener('open', () => resolve(new CDP(ws)));
      ws.addEventListener('error', () => reject(new Error('WS fail')));
      setTimeout(() => reject(new Error('WS timeout')), 3000);
    });
  } catch { cdp = null; }
}
if (!cdp) { console.log('起不了 headless Chrome'); process.exit(1); }
const evalIn = async (expr) => {
  const r = await cdp.send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: true });
  if (r.result && r.result.exceptionDetails) throw new Error('頁內例外: ' + String(r.result.exceptionDetails.exception?.description || r.result.exceptionDetails.text).slice(0, 300));
  return r.result && r.result.result ? r.result.result.value : null;
};
await cdp.send('Runtime.enable'); await cdp.send('Page.enable');
// u4k 教訓（11-25 診斷坐實）：headless 預設視埠 493px 高，真點擊座標會被
// Ribbon 固定列遮擋（elementFromPoint 命中 qword-tab-btn）→ caret 不動假紅。
// 探針必設高視埠，且點擊前斷言命中元素是畫布本身。
await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 1200, deviceScaleFactor: 1, mobile: false });
const goto = async (url) => { await cdp.send('Page.navigate', { url });
  for (let i = 0; i < 60; i++) { const s = await evalIn('document.readyState'); if (s === 'complete') break; await new Promise(r => setTimeout(r, 250)); }
  await new Promise(r => setTimeout(r, 300)); };
// 真鍵盤：CDP keyDown/keyUp（u4h 同口徑；Enter/Backspace/Delete 帶 windowsVirtualKeyCode）
const key = async (k, code, vk, opts = {}) => {
  const mods = opts.shift ? 8 : 0;
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: k, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers: mods });
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: k, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers: mods });
  await new Promise(r => setTimeout(r, 120)); // 等 refreshPages
};
const typeText = async (t) => { await cdp.send('Input.insertText', { text: t }); await new Promise(r => setTimeout(r, 150)); };

await goto(`${BASE}?sample=1`);
let settled = false;
for (let i = 0; i < 40 && !settled; i++) { settled = await evalIn('!!(window.QWordEngine && window.QWordEngine.engine && window.QWordEngine.isSettled && window.QWordEngine.isSettled())').catch(() => false); if (!settled) await new Promise(r => setTimeout(r, 400)); }
if (!settled) { console.log('引擎未定案，中止'); process.exit(1); }
pass('燈0 樣本開檔定案', `paras=${await evalIn('window.QWordEngine.paragraphCount()')}`);

// 佈點：游標放段 3 中點（真滑鼠點擊 canvas 反查），取 para/cp 基準
await evalIn(`document.querySelector('.qword-real-page[data-page-index="0"]')?.scrollIntoView({block:'start'}); true`);
await new Promise(r => setTimeout(r, 200));
const place = await evalIn(`(() => {
  const a = window.QWordEngine;
  const rect = a.caretRect(3, 12);
  const pageEl = document.querySelector('.qword-real-page[data-page-index="' + rect.page + '"]');
  pageEl.scrollIntoView({ block: 'center' });
  return rect.page;
})()`);
await new Promise(r => setTimeout(r, 350)); // 等滾動落定（CDP 點擊是視埠座標，卷動中會打空）
const placePt = await evalIn(`(() => {
  const a = window.QWordEngine;
  const rect = a.caretRect(3, 12);
  const pageEl = document.querySelector('.qword-real-page[data-page-index="' + rect.page + '"]');
  const canvas = pageEl.querySelector('.qword-page-canvas');
  const cr = canvas.getBoundingClientRect();
  const dl = a.getPageDisplayList(rect.page);
  const scale = cr.width / ((dl && dl.width_pt) || 595);
  const x = cr.left + rect.x_pt * scale, y = cr.top + rect.y_pt * scale;
  const hit = document.elementFromPoint(x, y);
  return { x, y, page: rect.page, hitCanvas: !!hit && hit.classList.contains('qword-page-canvas') };
})()`);
if (!placePt.hitCanvas) { fail('燈0a 點擊目標命中畫布', 'elementFromPoint 未命中 canvas（視埠遮擋？）'); }
else pass('燈0a 點擊目標命中畫布', `(${placePt.x.toFixed(0)},${placePt.y.toFixed(0)})`);
await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: placePt.x, y: placePt.y, button: 'left', clickCount: 1, buttons: 1 });
await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: placePt.x, y: placePt.y, button: 'left', clickCount: 1, buttons: 0 });
await new Promise(r => setTimeout(r, 200));
const before = await evalIn(`(() => ({ paras: window.QWordEngine.paragraphCount(), caret: window.QWordEngine.caret, text3: window.QWordEngine.paragraphText(3) }))()`);
if (!before.caret || before.caret.para !== 3) { fail('燈0b 真點擊落段 3', JSON.stringify(before.caret)); }
else pass('燈0b 真點擊落段 3', `cp=${before.caret.cp} paras=${before.paras}`);

// 燈1：Enter（硬換行）＝SplitParagraph：段數+1、caret 落新段起點
const cpBefore = before.caret.cp;
await key('Enter', 'Enter', 13);
const afterEnter = await evalIn(`(() => ({ paras: window.QWordEngine.paragraphCount(), caret: window.QWordEngine.caret, p3: window.QWordEngine.paragraphText(3), p4: window.QWordEngine.paragraphText(4) }))()`);
if (afterEnter.paras === before.paras + 1 && afterEnter.caret.para === 4 && afterEnter.caret.cp === 0) {
  pass('燈1 Enter 硬換行 SplitParagraph', `paras ${before.paras}→${afterEnter.paras}，caret=(4,0)，p3 長 ${cpBefore}`);
} else fail('燈1 Enter 硬換行 SplitParagraph', JSON.stringify({ before: before.paras, ...afterEnter, caret: afterEnter.caret }));

// 燈2：新段起點立刻打字（caret_after 無效→打字錯位的舊病復發鎖）
await typeText('甲');
const typedNew = await evalIn(`(() => ({ p4: window.QWordEngine.paragraphText(4), caret: window.QWordEngine.caret }))()`);
if (typedNew.p4.startsWith('甲')) pass('燈2 換行後新段打字落點正確', typedNew.p4.slice(0, 8));
else fail('燈2 換行後新段打字落點正確', JSON.stringify(typedNew));

// 燈3：段頭 Backspace＝退回縫合處（合併回段 3）
// 順序紀律：燈2 打了「甲」（caret=(4,1)），第一次 Backspace 正確刪「甲」落回
// (4,0)（Word 語義，非缺陷），第二次才觸發 merge。少這一步會拿「刪字」當
// 「合併」來斷言——11-25 探針自身假紅，記錄在此免再犯。
await key('Backspace', 'Backspace', 8);
const cleared = await evalIn(`(() => ({ caret: window.QWordEngine.caret, p4: window.QWordEngine.paragraphText(4) }))()`);
if (cleared.caret.cp !== 0 || cleared.p4.startsWith('甲')) {
  fail('燈3a 清場 Backspace 回段頭', JSON.stringify(cleared));
} else pass('燈3a 清場 Backspace 回段頭', `caret=(4,0)，p4 首字=${cleared.p4.slice(0, 6)}`);
await key('Backspace', 'Backspace', 8);
const afterMerge = await evalIn(`(() => ({ paras: window.QWordEngine.paragraphCount(), caret: window.QWordEngine.caret, p3: window.QWordEngine.paragraphText(3) }))()`);
if (afterMerge.paras === before.paras && afterMerge.caret.para === 3 && afterMerge.caret.cp === cpBefore
    && afterMerge.p3.startsWith(before.text3.slice(0, cpBefore))) {
  pass('燈3 段頭 Backspace 合併回縫合處', `paras→${afterMerge.paras}，caret=(3,${afterMerge.caret.cp})`);
} else fail('燈3 段頭 Backspace 合併回縫合處', JSON.stringify({ wantParas: before.paras, wantCaret: [3, cpBefore], ...afterMerge }));

// 燈4：Shift+ArrowRight 延伸選區＋打字替換選區（真鍵盤全路徑）
const selState = await evalIn(`(() => { const a = window.QWordEngine; a.caret = { para: 1, cp: 2 }; a.selection = null; return true; })()`);
await key('ArrowRight', 'ArrowRight', 39, { shift: true });
await key('ArrowRight', 'ArrowRight', 39, { shift: true });
const sel1 = await evalIn(`(() => { const s = window.QWordEngine.selection; return s && { a: s.start, b: s.end }; })()`);
if (sel1 && sel1.a.para === 1 && sel1.a.cp === 2 && sel1.b.cp === 4) pass('燈4a Shift+Right 延伸選區', `(${sel1.a.para},${sel1.a.cp})→(${sel1.b.para},${sel1.b.cp})`);
else fail('燈4a Shift+Right 延伸選區', JSON.stringify(sel1));
const textBefore = await evalIn('window.QWordEngine.paragraphText(1)');
await typeText('X');
const textAfter = await evalIn('window.QWordEngine.paragraphText(1)');
if (textAfter === textBefore.slice(0, 2) + 'X' + textBefore.slice(4)) pass('燈4b 打字替換選區', JSON.stringify(textAfter.slice(0, 8)));
else fail('燈4b 打字替換選區', `before=${JSON.stringify(textBefore.slice(0,8))} after=${JSON.stringify(textAfter.slice(0,8))}`);
await evalIn('window.QWordEngine.undo(); true');
await new Promise(r => setTimeout(r, 200));

// 燈5：Delete 鍵（選區→刪選區；游標→刪後一格）
await evalIn(`(() => { const a = window.QWordEngine; a.caret = { para: 5, cp: 3 }; a.selection = null; return true; })()`);
const t5 = await evalIn('window.QWordEngine.paragraphText(5)');
await key('Delete', 'Delete', 46);
const t5b = await evalIn('window.QWordEngine.paragraphText(5)');
if (t5b === t5.slice(0, 3) + t5.slice(4)) pass('燈5 Delete 刪游標後一格', `長 ${t5.length}→${t5b.length}`);
else fail('燈5 Delete 刪游標後一格', `before=${t5.slice(0,6)} after=${t5b.slice(0,6)}`);

// 燈6：跨段選區（起 p1 終 p2）打字＝跨段合併刪除後插入（K29b 引擎能力面）
await evalIn(`(() => { const a = window.QWordEngine; a.selection = { target: 'Body', start: { para: 1, cp: 6 }, end: { para: 2, cp: 3 } }; a.caret = { para: 2, cp: 3 }; return true; })()`);
const pc6 = await evalIn('window.QWordEngine.paragraphCount()');
await typeText('Y');
const st6 = await evalIn(`(() => ({ paras: window.QWordEngine.paragraphCount(), p1: window.QWordEngine.paragraphText(1) }))()`);
if (st6.paras === pc6 - 1 && st6.p1.includes('Y')) pass('燈6 跨段選區打字吞段', `paras ${pc6}→${st6.paras}`);
else fail('燈6 跨段選區打字吞段', JSON.stringify(st6));

// 燈7：反向選區（終點在起點前）不崩——打字仍按正規化序執行
await evalIn(`(() => { const a = window.QWordEngine; a.selection = { target: 'Body', start: { para: 3, cp: 8 }, end: { para: 3, cp: 4 } }; a.caret = { para: 3, cp: 4 }; return true; })()`);
const ok7 = await evalIn(`(() => { try { const r = window.QWordEngine.typeAt(undefined, undefined, 'Z'); return r.ok === true; } catch (e) { return 'throw:' + e.message; } })()`);
if (ok7 === true) pass('燈7 反向選區正規化', 'typeAt ok');
else fail('燈7 反向選區正規化', String(ok7));

// 燈8：段 0 起點 Backspace＝誠實拒絕（不崩、不動文檔）
await evalIn(`(() => { const a = window.QWordEngine; a.selection = null; a.caret = { para: 0, cp: 0 }; return true; })()`);
const paras8 = await evalIn('window.QWordEngine.paragraphCount()');
await key('Backspace', 'Backspace', 8);
const paras8b = await evalIn('window.QWordEngine.paragraphCount()');
if (paras8b === paras8) pass('燈8 文件開頭 Backspace 誠實拒絕', `paras 恆 ${paras8}`);
else fail('燈8 文件開頭 Backspace 誠實拒絕', `paras ${paras8}→${paras8b}`);

fs.mkdirSync(path.dirname(OUT), { recursive: true });
fs.writeFileSync(OUT, JSON.stringify({ at: new Date().toISOString(), base: BASE, results }, null, 2));
try { const shot = await cdp.send('Page.captureScreenshot', { format: 'png' });
  if (shot && shot.data) fs.writeFileSync(OUT.replace(/\.json$/, '.png'), Buffer.from(shot.data, 'base64')); } catch {}
ws.close(); chromeChild.kill('SIGKILL');
const fails = results.filter(r => !r.ok);
console.log(`\n判定：${fails.length === 0 ? 'PASS' : 'FAIL'}（${fails.length} 項紅）`);
process.exit(fails.length === 0 ? 0 : 1);
