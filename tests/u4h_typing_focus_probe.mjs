// U4h 打字焦點閉環探針（C1 回歸鎖）
// 根因實錄：真實滑鼠點擊下，相容 mousedown 的預設行為把 pointerdown handler
// 剛給的 inputSink 焦點撤回收 body → 真鍵盤事件落空 → 「有游標但無法打字」。
// 本探針鎖的是**真實輸入路徑**（CDP Input.dispatchMouseEvent/KeyEvent，非
// 程式化 focus）：真點擊後焦點必須留在輸入錨點、真敲鍵必須寫進段落。
// u4c 那類程式化路徑探針對這條缺陷無鑑別力，故必須有本卡。
//
// 用法：node tests/u4h_typing_focus_probe.mjs [editor_url]
//   預帶 http://127.0.0.1:8791/word.html；自起 headless Chrome（獨立 profile
//   ＋專用埠 9224，沿用本倉 u4f 探針自起慣例，不借用戶瀏覽器）。
// 退出碼：0 全綠；1 任一燈紅（逐盞點名）。
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const BASE_URL = process.env.U4H_URL || process.argv[2] || 'http://127.0.0.1:8791/word.html';
const CHROME = process.env.CHROME_BIN || '/usr/bin/google-chrome';
const CDP_PORT = process.env.U4H_CDP_PORT || '9224';
const CDP_HTTP = `http://127.0.0.1:${CDP_PORT}`;
const OUT = process.env.U4H_OUT || path.join(ROOT, 'evidence/word-ui-u4h-typing-evidence.json');

class CDP {
  constructor(ws) {
    this.ws = ws; this.id = 0; this.pending = new Map(); this.events = [];
    ws.addEventListener('message', (m) => {
      const d = JSON.parse(m.data);
      if (d.id && this.pending.has(d.id)) { this.pending.get(d.id)(d); this.pending.delete(d.id); }
      else if (d.method) this.events.push(d.method);
    });
  }
  send(method, params = {}, sessionId) {
    return new Promise((resolve, reject) => {
      const i = ++this.id;
      this.pending.set(i, resolve);
      this.ws.send(JSON.stringify({ id: i, method, params, sessionId }));
      setTimeout(() => { if (this.pending.has(i)) { this.pending.delete(i); reject(new Error(`CDP 逾時: ${method}`)); } }, 20000);
    });
  }
}

const results = [];
const fail = (name, detail) => { results.push({ name, ok: false, detail }); console.log(`  ✗ ${name}｜${detail}`); };
const pass = (name, detail) => { results.push({ name, ok: true, detail }); console.log(`  ✓ ${name}｜${detail}`); };

const created = null; // 自起 Chrome（下方），不用用戶瀏覽器的 tab
const profileDir = fs.mkdtempSync(path.join(os.tmpdir(), 'u4h-probe-'));
const chromeChild = spawn(CHROME, ['--headless=new', `--user-data-dir=${profileDir}`, `--remote-debugging-port=${CDP_PORT}`,
  '--no-first-run', '--no-default-browser-check', '--disable-gpu', 'about:blank'], { stdio: 'ignore' });
let ws = null;
let cdp = null;
for (let i = 0; i < 40 && !cdp; i++) {
  await new Promise((r) => setTimeout(r, 400));
  try {
    const list = await (await fetch(`${CDP_HTTP}/json/list`)).json();
    const t = list.find((x) => x.type === 'page');
    if (!t) continue;
    ws = new WebSocket(t.webSocketDebuggerUrl);
    cdp = await new Promise((resolve, reject) => {
      ws.addEventListener('open', () => resolve(new CDP(ws)));
      ws.addEventListener('error', () => reject(new Error('CDP WebSocket 連線失敗')));
      setTimeout(() => reject(new Error('CDP open 逾時')), 3000);
    });
  } catch { cdp = null; }
}
if (!cdp) { console.log('起不了 headless Chrome'); process.exit(1); }
const evalIn = async (expr) => {
  const r = await cdp.send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: true });
  if (r.error) throw new Error('CDP 錯誤: ' + JSON.stringify(r.error).slice(0, 200));
  const res = r.result && r.result.result;
  if (r.result && r.result.exceptionDetails) throw new Error('頁內例外: ' + String(r.result.exceptionDetails.exception?.description || r.result.exceptionDetails.text).slice(0, 300));
  return res ? res.value : null;
};

await cdp.send('Runtime.enable');
await cdp.send('Page.enable');
await cdp.send('Page.navigate', { url: BASE_URL });

let ready = false;
let lastErr = null;
for (let i = 0; i < 40 && !ready; i++) {
  ready = await evalIn('!!(window.QWordShell && window.QWordEngine)').catch((e) => { lastErr = String(e).slice(0, 250); return false; });
  if (!ready) await new Promise((r) => setTimeout(r, 400));
}
if (!ready) {
  console.log('編輯器未就緒' + (lastErr ? '｜' + lastErr : ''));
  try {
    console.log('DIAG:', JSON.stringify(await evalIn('(() => ({ href: location.href, title: document.title, scripts: [].map.call(document.scripts, s => s.src.split("/").pop()), err: window.__lastErr || null }))()')));
  } catch (e) { console.log('DIAG 也掛:', String(e).slice(0, 150)); }
  try { ws.close(); } catch {}
  chromeChild.kill('SIGKILL');
  process.exit(1);
}
await evalIn('window.QWordShell.openDefaultSample()');
for (let i = 0; i < 40; i++) {
  const s = await evalIn('!!(window.QWordEngine.engine && window.QWordEngine.isSettled && window.QWordEngine.isSettled())').catch(() => false);
  if (s) break;
  await new Promise((r) => setTimeout(r, 300));
}

// 選正文段
const target = await evalIn(`(() => {
  const E = window.QWordEngine;
  for (let p = 0; p < E.paragraphCount(); p++) {
    const t = E.paragraphText(p) || '';
    if (Array.from(t).length > 10) { window.__u4hPara = p; return { p, text: t }; }
  }
  return null;
})()`);
if (!target) { console.log('找不到正文段'); process.exit(1); }

// 幾何：caret_rect → 視口座標（用 y 中間偏上，避開行緣）
const pt = await evalIn(`(() => {
  const E = window.QWordEngine;
  const r = E.caretRect(window.__u4hPara, 5);
  if (!r) return null;
  const wrapper = document.querySelectorAll('.qword-real-page')[r.page];
  const canvas = wrapper.querySelector('canvas');
  const rect = canvas.getBoundingClientRect();
  const dl = E.getPageDisplayList(r.page);
  const scale = canvas.width / (dl.width_pt || 595);
  return { x: rect.left + r.x_pt * scale, y: rect.top + (r.y_pt - 4) * scale };
})()`);
if (!pt) { console.log('caretRect 回 None'); process.exit(1); }

// 燈1：真實點擊（CDP mouse 事件＝含相容 mouse 語義的完整輸入路徑）
await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: pt.x, y: pt.y, button: 'left', clickCount: 1, buttons: 1 });
await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: pt.x, y: pt.y, button: 'left', clickCount: 1, buttons: 0 });
await new Promise((r) => setTimeout(r, 400));
const focusAfterClick = await evalIn('document.activeElement && document.activeElement.id');
const caretAfterClick = await evalIn('JSON.stringify(window.QWordEngine.caret)');
if (focusAfterClick === 'qword-input-sink') pass('燈1 真點擊後焦點留在輸入錨點', `activeElement=${focusAfterClick} caret=${caretAfterClick}`);
else fail('燈1 真點擊後焦點留在輸入錨點', `activeElement=${focusAfterClick}（C1 缺陷形態：相容 mousedown 預設把焦點收回 body）`);

// 燈2：真鍵盤打字寫進段落（非程式化灌值）
const TYPED = 'u4hok';
for (const ch of TYPED) {
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', text: ch, key: ch, code: 'Key' + ch.toUpperCase(), windowsVirtualKeyCode: ch.toUpperCase().charCodeAt(0) });
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: ch, code: 'Key' + ch.toUpperCase(), windowsVirtualKeyCode: ch.toUpperCase().charCodeAt(0) });
  await new Promise((r) => setTimeout(r, 80));
}
await new Promise((r) => setTimeout(r, 700));
const typedOk = await evalIn(`(window.QWordEngine.paragraphText(window.__u4hPara) || '').includes('${TYPED}')`);
if (typedOk) pass('燈2 真鍵盤打字進段', `段落 ${target.p} 已含 "${TYPED}"`);
else fail('燈2 真鍵盤打字進段', `段落 ${target.p} 未見 "${TYPED}"（打字鏈斷點）`);

// 燈3：打完字 focus 仍在（打字後焦点不被搶）＋ caret 推進
const afterState = await evalIn(`(() => ({
  focus: document.activeElement && document.activeElement.id,
  caret: window.QWordEngine.caret,
  dirty: window.QWordEngine.isDirty
}))()`);
if (afterState.focus === 'qword-input-sink' && afterState.dirty) pass('燈3 打字後焦點持續＋髒旗標', JSON.stringify(afterState));
else fail('燈3 打字後焦點持續＋髒旗標', JSON.stringify(afterState));

fs.mkdirSync(path.dirname(OUT), { recursive: true });
fs.writeFileSync(OUT, JSON.stringify({ at: new Date().toISOString(), baseUrl: BASE_URL, target, results }, null, 2));
ws.close();
chromeChild.kill('SIGKILL');

const fails = results.filter((r) => !r.ok);
console.log(`\n判定：${fails.length === 0 ? 'PASS' : 'FAIL'}（${fails.length} 項紅）`);
process.exit(fails.length === 0 ? 0 : 1);
