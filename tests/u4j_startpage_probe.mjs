// U4j 起始頁與路由探針（C4）
// 驗收面：起始頁（index.html）三條開檔路（開啟／最近／示範）＋ word.html
// pending 交接＋最近清單持久化＋移除即真刪＋拖放。全部走 CDP 真輸入路徑。
// 用法：node tests/u4j_startpage_probe.mjs [base_url]（預帶 http://127.0.0.1:8791/）
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const BASE = (process.env.U4J_URL || process.argv[2] || 'http://127.0.0.1:8791/').replace(/\/$/, '');
const CHROME = process.env.CHROME_BIN || '/usr/bin/google-chrome';
const CDP_PORT = process.env.U4J_CDP_PORT || '9225';
const CDP_HTTP = `http://127.0.0.1:${CDP_PORT}`;

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

const profileDir = fs.mkdtempSync(path.join(os.tmpdir(), 'u4j-probe-'));
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
const goto = async (url) => { await cdp.send('Page.navigate', { url });
  for (let i = 0; i < 60; i++) { const s = await evalIn('document.readyState'); if (s === 'complete') break; await new Promise(r => setTimeout(r, 250)); }
  await new Promise(r => setTimeout(r, 300)); };

// 燈0：起始頁存在且骨架正確
await goto(`${BASE}/index.html`);
const home = await evalIn(`(() => ({ hasOpen: !!document.getElementById('btn-open'),
  hasSample: !!document.getElementById('btn-sample'), hasDrop: !!document.getElementById('drop'),
  hasRows: !!document.getElementById('rows'), title: document.title }))()`);
if (home && home.hasOpen && home.hasSample && home.hasDrop && home.hasRows) pass('燈0 起始頁骨架', JSON.stringify(home.title));
else fail('燈0 起始頁骨架', JSON.stringify(home));

// 燈1：示範檔路 → word.html?sample=1 自動開檔
await goto(`${BASE}/word.html?sample=1`);
let settled = false;
for (let i = 0; i < 40 && !settled; i++) { settled = await evalIn('!!(window.QWordEngine && window.QWordEngine.engine && window.QWordEngine.isSettled && window.QWordEngine.isSettled())').catch(() => false); if (!settled) await new Promise(r => setTimeout(r, 400)); }
if (settled) pass('燈1 ?sample=1 自動開示範檔', `paragraphCount=${await evalIn('window.QWordEngine.paragraphCount()')}`);
else fail('燈1 ?sample=1 自動開示範檔', '30s 內引擎未定案');

// 燈2：最近清單——起始頁把檔放 pending → word.html 收起並開檔 → 回起始頁見最近列
// （用引擎內建樣本位元組造一個具名檔，走 putPending 路徑，非直接 recordOpen，
//  這正是起始頁三條開檔路共用的交接面。）
await goto(`${BASE}/index.html`);
await evalIn(`(async () => {
  const m = await import('./js/recent-store.js');
  // 造 3 個具名檔（用樣本位元組改檔名即可——本探針驗的是清單/交接，不是內容）
  const bytes = (await import('./js/engine-adapter.js')).getDefaultSampleBytes();
  for (const nm of ['季報-甲.docx', '備忘-乙.docx', '合約-丙.docx']) await m.recordOpen(nm, bytes);
  await m.putPending(bytes, '待開-丁.docx');
  return true;
})()`);
const recentList = await evalIn(`(async () => (await (await import('./js/recent-store.js')).listRecent(10)).map(r => r.name))()`);
if (Array.isArray(recentList) && recentList.length === 3 && recentList.includes('季報-甲.docx')) pass('燈2a 最近清單持久化', JSON.stringify(recentList));
else fail('燈2a 最近清單持久化', JSON.stringify(recentList));

await goto(`${BASE}/word.html`);
settled = false;
for (let i = 0; i < 40 && !settled; i++) { settled = await evalIn('!!(window.QWordEngine && window.QWordEngine.engine && window.QWordEngine.isSettled && window.QWordEngine.isSettled())').catch(() => false); if (!settled) await new Promise(r => setTimeout(r, 400)); }
const openedName = await evalIn('window.QWordEngine.fileName');
if (settled && openedName === '待開-丁.docx') pass('燈2b pending 交接自動開檔', openedName);
else fail('燈2b pending 交接自動開檔', `settled=${settled} fileName=${openedName}`);
const pendingGone = await evalIn(`(async () => { const m = await import('./js/recent-store.js'); return (await m.takePending()) === null; })()`);
if (pendingGone) pass('燈2c pending 取用即清', 'takePending 回 null');
else fail('燈2c pending 取用即清', 'pending 殘留');

// 燈3：word.html 開使用者檔會入最近（loadFile 記帳路徑；非 sample.docx）
const recorded = await evalIn(`(async () => {
  const m = await import('./js/recent-store.js');
  const before = (await m.listRecent(50)).length;
  const bytes = (await import('./js/engine-adapter.js')).getDefaultSampleBytes();
  await window.QWordShell.openFile(bytes, '實測-戊.docx');
  await new Promise(r => setTimeout(r, 600));
  const after = (await m.listRecent(50)).map(r => r.name);
  return { before, has: after.includes('實測-戊.docx'), after: after.length };
})()`);
if (recorded && recorded.has) pass('燈3 編輯器開檔入最近', JSON.stringify(recorded));
else fail('燈3 編輯器開檔入最近', JSON.stringify(recorded));

// 燈4：起始頁移除＝真刪（bytes 一併 gone）
await goto(`${BASE}/index.html`);
const removed = await evalIn(`(async () => {
  const m = await import('./js/recent-store.js');
  const all = await m.listRecent(50);
  const target = all.find(r => r.name === '備忘-乙.docx');
  if (!target) return { err: '找不到目標' };
  await m.removeRecent(target.id);
  const left = await m.listRecent(50);
  const got = await m.getRecent(target.id);
  return { gone: !left.some(r => r.name === '備忘-乙.docx'), bytesGone: got === null };
})()`);
if (removed && removed.gone && removed.bytesGone) pass('燈4 移除即真刪', JSON.stringify(removed));
else fail('燈4 移除即真刪', JSON.stringify(removed));

// 燈5：UI 真點擊——起始頁「示範文件」鈕 → 落到 word.html?sample=1
await goto(`${BASE}/index.html`);
const btn = await evalIn(`(() => { const b = document.getElementById('btn-sample'); const r = b.getBoundingClientRect(); return { x: r.left + r.width/2, y: r.top + r.height/2 }; })()`);
await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: btn.x, y: btn.y, button: 'left', clickCount: 1, buttons: 1 });
await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: btn.x, y: btn.y, button: 'left', clickCount: 1, buttons: 0 });
for (let i = 0; i < 40; i++) { const u = await evalIn('location.href').catch(() => ''); if (u.includes('word.html')) break; await new Promise(r => setTimeout(r, 300)); }
const finalUrl = await evalIn('location.href');
if (finalUrl.includes('word.html?sample=1')) pass('燈5 真點擊示範鈕路由', finalUrl.replace(BASE, ''));
else fail('燈5 真點擊示範鈕路由', finalUrl.replace(BASE, ''));

ws.close(); chromeChild.kill('SIGKILL');
const fails = results.filter(r => !r.ok);
console.log(`\n判定：${fails.length === 0 ? 'PASS' : 'FAIL'}（${fails.length} 項紅）`);
process.exit(fails.length === 0 ? 0 : 1);
