/**
 * u6_mac_ime_probe.mjs — macOS 繁中輸入法三缺陷機械判據（U6-MAC-IME，先紅後綠）
 *
 * 用戶實錄（2026-11-27，Mac Safari/Chrome 繁中鍵盤）：
 *   (1) 打字後不會即時顯示；(2) IME 選字確認按 Enter **必定換行**（每個詞各占一行）；
 *   (3) 中文輸入時 cursor 與畫面實際落點不一致。
 *
 * 本卡鎖五盞燈（CDP 模擬 macOS 注音/倉頡事件序）：
 *   L1 Enter 分段鎖 —— 合成窗開著 dispatch keydown Enter → 段落數不得 +1，定稿字入該段。
 *   L2 合成期 Backspace 鎖 —— 合成窗開著 dispatch Backspace → 引擎正文零變動（鍵歸 IME）。
 *   L3 候選窗錨點鎖 —— inputSink 必錨在畫面游標處（≤2px），不得是 -9999 離屏常數。
 *   L4 合成回顯鎖 —— compositionupdate 期間 DOM 有合成 overlay 含拼音，引擎與正文零拼音；
 *                    compositionend 後 overlay 消失、定稿字入正文。
 *   L5 入文點快照鎖 —— 合成窗開著時以另一途徑移動 caret → 定稿仍落 compositionstart 快照點。
 *
 * 事件序忠實度（CDP 協議實查 /json/protocol 坐實，非臆造）：
 *   `Input.imeSetComposition` → 於 focus 元素派發 compositionstart＋compositionupdate；
 *   `Input.insertText`        → 派發 input＋compositionend(data=定稿串)；
 *   `Input.dispatchKeyEvent`  → **協議無 isComposing 參數**（參數只有 type/modifiers/text/key/
 *                               code/windowsVirtualKeyCode…），合成期按鍵的 isComposing 位由
 *                               瀏覽器依 IME 合成態自行標記，探針端無法直接指定。故本卡修法
 *                               以**外殼自持旗標 `this._composing` 為主防線**（PM 定案軸線），
 *                               並由 L0 前提自證燈坐實「合成窗真的開著」（compositionstart
 *                               實錄到位才進 L1／L2／L5），避免事件序沒送成卻假綠。
 *
 * 回塞鑑別：`U6_IME_BUGGY=1..5`（逗號可並選／all）經 Page.addScriptToEvaluateOnNewDocument
 *   注入 `window.process.env`（先於頁內任何腳本），生產碼**於事件時求值**（非模組層快照）。
 *   每盞燈各有對應回塞態必紅；探針入口先 delete process.env.U6_IME_BUGGY 防外洩。
 *
 * 用法：node tests/u6_mac_ime_probe.mjs
 *   env：U6_URL（預設自起本 worktree 靜態站）、U6_CDP_PORT（預設 9226）、
 *        U6_DPR（逗號表，預設 '1,2'）、U6_BUGGY（等於 U6_IME_BUGGY 的別名）、U6_OUT 證據檔。
 * 退出碼：0 全綠；1 任一燈紅（逐盞點名）。
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
// 防外洩：本程序自身（Node）不帶 BUGGY 旗標進子程序／證據檔
const BUGGY_RAW = process.env.U6_IME_BUGGY || process.env.U6_BUGGY || '';
delete process.env.U6_IME_BUGGY;
delete process.env.U6_BUGGY;
const BUGGY = BUGGY_RAW.trim();

const CDP_BIN = process.env.CHROME_BIN || '/usr/bin/google-chrome';
const CDP_PORT = process.env.U6_CDP_PORT || '9226';
const CDP_HTTP = `http://127.0.0.1:${CDP_PORT}`;
const HTTP_PORT = process.env.U6_HTTP_PORT || '8931';
const DPRS = (process.env.U6_DPR || '1,2').split(',').map((x) => Number(x.trim())).filter((x) => x > 0);
const OUT = process.env.U6_OUT || path.join(ROOT, 'evidence', `word-ui-u6-mac-ime-evidence${BUGGY ? `.buggy${BUGGY}` : ''}.json`);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** 靜態站必自起且 serve **本 worktree**：8791 是主倉（main tip），沿用＝拿舊碼測，假綠源。 */
const httpChild = spawn('python3', ['-m', 'http.server', HTTP_PORT, '--bind', '127.0.0.1'], {
  cwd: ROOT, stdio: 'ignore', detached: true,
});
const BASE_URL = process.env.U6_URL || `http://127.0.0.1:${HTTP_PORT}/word.html`;

class CDP {
  constructor(ws) {
    this.ws = ws; this.id = 0; this.pending = new Map();
    ws.addEventListener('message', (m) => {
      const d = JSON.parse(m.data);
      if (d.id && this.pending.has(d.id)) { this.pending.get(d.id)(d); this.pending.delete(d.id); }
    });
  }
  send(method, params = {}) {
    return new Promise((resolve, reject) => {
      const i = ++this.id;
      this.pending.set(i, resolve);
      this.ws.send(JSON.stringify({ id: i, method, params }));
      setTimeout(() => { if (this.pending.has(i)) { this.pending.delete(i); reject(new Error(`CDP 逾時: ${method}`)); } }, 25000);
    });
  }
}

const results = [];
const fail = (name, detail) => { results.push({ name, ok: false, detail }); console.log(`  ✗ ${name}｜${detail}`); };
const pass = (name, detail) => { results.push({ name, ok: true, detail }); console.log(`  ✓ ${name}｜${detail}`); };

// ---------- 起環境（headless Chrome 自起，獨立 profile，不借用戶瀏覽器）----------
const profileDir = fs.mkdtempSync(path.join(os.tmpdir(), 'u6-ime-probe-'));
const chromeChild = spawn(CDP_BIN, ['--headless=new', `--user-data-dir=${profileDir}`, `--remote-debugging-port=${CDP_PORT}`,
  '--no-first-run', '--no-default-browser-check', '--disable-gpu', '--hide-scrollbars', 'about:blank'], { stdio: 'ignore' });

let cdp = null, ws = null;
for (let i = 0; i < 50 && !cdp; i++) {
  await sleep(400);
  try {
    const list = await (await fetch(`${CDP_HTTP}/json/list`)).json();
    const t = list.find((x) => x.type === 'page');
    if (!t) continue;
    ws = new WebSocket(t.webSocketDebuggerUrl);
    cdp = await new Promise((resolve, reject) => {
      ws.addEventListener('open', () => resolve(new CDP(ws)));
      ws.addEventListener('error', () => reject(new Error('CDP WebSocket 連線失敗')));
      setTimeout(() => reject(new Error('CDP open 逾時')), 4000);
    });
  } catch { cdp = null; }
}
if (!cdp) { console.log('起不了 headless Chrome'); try { process.kill(-httpChild.pid, 'SIGKILL'); } catch {} chromeChild.kill('SIGKILL'); process.exit(1); }

const evalIn = async (expr) => {
  const r = await cdp.send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: true, timeout: 120000 });
  if (r.error) throw new Error('CDP 錯誤: ' + JSON.stringify(r.error).slice(0, 200));
  if (r.result && r.result.exceptionDetails) throw new Error('頁內例外: ' + String(r.result.exceptionDetails.exception?.description || r.result.exceptionDetails.text).slice(0, 300));
  return r.result && r.result.result ? r.result.result.value : null;
};

await cdp.send('Runtime.enable');
await cdp.send('Page.enable');
await cdp.send('Network.enable');
await cdp.send('Network.setCacheDisabled', { cacheDisabled: true });
// 回塞鎖注入：addScriptToEvaluateOnNewDocument 先於頁內任何腳本執行（時點保證，
// 且生產碼仍採「事件時求值」雙保險——注入早於/晚於模組求值都咬得住）。
if (BUGGY) {
  await cdp.send('Page.addScriptToEvaluateOnNewDocument', {
    source: `window.process = { env: Object.assign({}, window.process && window.process.env, { U6_IME_BUGGY: ${JSON.stringify(BUGGY)} }) };`,
  });
}

const pageErrors = [];
cdp.ws.addEventListener('message', (m) => {
  const d = JSON.parse(m.data);
  if (d.method === 'Runtime.exceptionThrown') pageErrors.push(String(d.params?.exceptionDetails?.exception?.description || d.params?.exceptionDetails?.text || '').slice(0, 160));
});

async function boot(dpr) {
  await cdp.send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 1400, deviceScaleFactor: dpr, mobile: false });
  await cdp.send('Page.navigate', { url: BASE_URL });
  for (let i = 0; i < 80; i++) { if (await evalIn('document.readyState') === 'complete') break; await sleep(200); }
  await evalIn('window.QWordShell.openDefaultSample()');
  for (let i = 0; i < 120; i++) {
    const ok = await evalIn('!!(window.QWordEngine.engine && window.QWordEngine.paragraphCount() > 0 && window.QWordEngine.isSettled())').catch(() => false);
    if (ok) break;
    await sleep(200);
  }
  await sleep(600);
  // 事件軌跡＋合成窗自證面（capture 期，零副作用）
  await evalIn(`(function(){
    window.__u6trace = [];
    const sink = document.getElementById('qword-input-sink');
    ['compositionstart','compositionupdate','compositionend','input','keydown','blur','focus'].forEach(function(t){
      sink.addEventListener(t, function(e){
        window.__u6trace.push({ t: t, key: e.key || null, data: (e.data != null ? String(e.data).slice(0,12) : null), comp: !!e.isComposing, val: String(sink.value).slice(0,12) });
      }, true);
    });
    document.addEventListener('keydown', function(e){ window.__u6trace.push({ t: 'doc-keydown', key: e.key, comp: !!e.isComposing }); }, true);
  })()`);
}

/** 真滑鼠點擊段內 cp 處：同時完成「caret 定位＋焦點留錨點」（真輸入路徑，非程式化 focus） */
async function clickIntoCaret(para, cp) {
  const geo = await evalIn(`(async function(){
    const A = window.QWordEngine;
    A.selection = null;
    const r = A.caretRect(${para}, ${cp});
    if (!r) return null;
    const w = document.querySelectorAll('.qword-real-page')[r.page];
    w.scrollIntoView({ block: 'center' });
    await new Promise(function(x){ setTimeout(x, 260); });
    const c = w.querySelector('canvas');
    const rect = c.getBoundingClientRect();
    const dl = A.getPageDisplayList(r.page);
    const scale = rect.width / ((dl && dl.width_pt) || 595);
    return { x: Math.round(rect.left + (r.x_pt + 1) * scale), y: Math.round(rect.top + (r.y_pt + (r.height_pt || 12) / 2) * scale), page: r.page };
  })()`);
  if (!geo) return null;
  await cdp.send('Input.dispatchMouseEvent', { type: 'mousePressed', x: geo.x, y: geo.y, button: 'left', clickCount: 1, buttons: 1 });
  await cdp.send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: geo.x, y: geo.y, button: 'left', clickCount: 1, buttons: 0 });
  await sleep(320);
  return geo;
}

const dump = () => evalIn(`JSON.stringify({ trace: window.__u6trace, sink: document.getElementById('qword-input-sink').value, caret: window.QWordEngine.caret, paras: (() => { const A = window.QWordEngine; const o = []; for (let p = 0; p < A.paragraphCount(); p++) o.push(A.paragraphText(p) || ''); return o; })() })`);
const clearTrace = () => evalIn('window.__u6trace = []; true');

async function runRound(dpr) {
  console.log(`\n===== DPR=${dpr}（桌面 viewport 1280x1400）${BUGGY ? `｜回塞態 U6_IME_BUGGY=${BUGGY}` : '｜真樹態'} =====`);
  await boot(dpr);
  const tag = `DPR${dpr}`;

  // ---------- L0 前提自證：真點擊後焦點留錨點＋合成窗事件序送得成 ----------
  const P = await evalIn(`(() => { const A = window.QWordEngine; for (let p = 0; p < A.paragraphCount(); p++) { if (Array.from(A.paragraphText(p)||'').length > 12) return p; } return -1; })()`);
  if (P < 0) { fail(`${tag}/L0 找到正文段`, 'sample 無 >12 字元段落，考卷無法出題'); return; }
  const before0 = JSON.parse(await dump());
  await clickIntoCaret(P, 3);
  const focusId = await evalIn('document.activeElement && document.activeElement.id');
  await clearTrace();
  await cdp.send('Input.imeSetComposition', { text: 'zhong', selectionStart: 5, selectionEnd: 5 });
  await sleep(320);
  const tr0 = JSON.parse(await dump()).trace;
  const sawStart = tr0.some((x) => x.t === 'compositionstart');
  const sawUpdate = tr0.some((x) => x.t === 'compositionupdate');
  if (focusId === 'qword-input-sink' && sawStart) pass(`${tag}/L0 前提自證`, `focus=inputSink，compositionstart${sawUpdate ? '＋compositionupdate' : ''} 實錄到位（事件序成立，後續燈非空轉）`);
  else fail(`${tag}/L0 前提自證`, `focus=${focusId} compositionstart=${sawStart} update=${sawUpdate}——事件序未成立，本輪燈不可信`);
  // 收掉這輪合成窗（以空字串取消＝IME 慣例）
  await cdp.send('Input.imeSetComposition', { text: '', selectionStart: 0, selectionEnd: 0 });
  await sleep(200);
  await cdp.send('Input.insertText', { text: '' });
  await sleep(200);

  // ---------- L1 Enter 分段鎖 ----------
  const len0 = await evalIn(`Array.from(window.QWordEngine.paragraphText(${P})||'').length`);
  const pc0 = await evalIn('window.QWordEngine.paragraphCount()');
  await clickIntoCaret(P, len0);                       // caret → 該段尾
  await clearTrace();
  await cdp.send('Input.imeSetComposition', { text: 'danbing', selectionStart: 7, selectionEnd: 7 });
  await sleep(280);
  const winOpen1 = JSON.parse(await dump()).trace.some((x) => x.t === 'compositionstart');
  // macOS 選字確認：Enter 的 keydown 照樣進外殼（合成窗開著）
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 36, text: '\r', unmodifiedText: '\r' });
  await sleep(160);
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 36 });
  await sleep(160);
  await cdp.send('Input.insertText', { text: '蛋餅' });
  await sleep(520);
  const s1 = JSON.parse(await dump());
  const pc1 = s1.paras.length;
  const grewTail = (s1.paras[P] || '').endsWith('蛋餅');
  const pinyinLeak = s1.paras.some((t) => String(t).includes('danbing'));
  if (!winOpen1) fail(`${tag}/L1 Enter 分段鎖`, '合成窗未開（compositionstart 缺席）——測試無效');
  else if (pc1 !== pc0) fail(`${tag}/L1 Enter 分段鎖`, `段落數 ${pc0}→${pc1}（IME commit 的 Enter 被當分段＝用戶「每詞一行」根因）｜trace=${JSON.stringify(s1.trace.slice(-6))}`);
  else if (!grewTail) fail(`${tag}/L1 Enter 分段鎖`, `段 ${P} 尾未見「蛋餅」：${JSON.stringify((s1.paras[P] || '').slice(-8))}`);
  else if (pinyinLeak) fail(`${tag}/L1 Enter 分段鎖`, '拼音殘字混入正文（合成期寫入引擎）');
  else pass(`${tag}/L1 Enter 分段鎖`, `段落數恆 ${pc1}、段 ${P} 尾追加「蛋餅」、全文無拼音殘字`);

  // ---------- L2 合成期 Backspace 鎖 ----------
  const snap2 = JSON.parse(await dump());
  const paras2 = snap2.paras.join('\u0001');
  await clickIntoCaret(P, 3);
  await clearTrace();
  await cdp.send('Input.imeSetComposition', { text: 'cangjie', selectionStart: 7, selectionEnd: 7 });
  await sleep(280);
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'Backspace', code: 'Backspace', windowsVirtualKeyCode: 8, nativeVirtualKeyCode: 51 });
  await sleep(150);
  await cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'Backspace', code: 'Backspace', windowsVirtualKeyCode: 8, nativeVirtualKeyCode: 51 });
  await sleep(420);
  const s2 = JSON.parse(await dump());
  const winOpen2 = s2.trace.some((x) => x.t === 'compositionstart');
  const engineSame = s2.paras.join('\u0001') === paras2;
  await cdp.send('Input.insertText', { text: '' });   // 取消合成，還原環境
  await sleep(300);
  const s2b = JSON.parse(await dump());
  if (!winOpen2) fail(`${tag}/L2 合成期 Backspace 鎖`, '合成窗未開——測試無效');
  else if (!engineSame) fail(`${tag}/L2 合成期 Backspace 鎖`, '合成窗開著按 Backspace，引擎正文被改動（外殼與 IME 搶鍵）');
  else if (s2b.paras.join('\u0001') !== paras2) fail(`${tag}/L2 合成期 Backspace 鎖`, `取消合成後正文仍變（殘留寫入）：${JSON.stringify(s2b.paras[P])}`);
  else pass(`${tag}/L2 合成期 Backspace 鎖`, '合成窗開著按 Backspace → 引擎正文零變動（鍵完整歸 IME）');

  // ---------- L3 候選窗錨點鎖 ----------
  await clickIntoCaret(P, 5);
  await sleep(200);
  await evalIn(`(() => { const A = window.QWordEngine; A.caret = { para: ${P}, cp: 7 }; window.QWordShell.render(); return true; })()`);
  await sleep(260);
  const g3 = JSON.parse(await evalIn(`(() => {
    const sink = document.getElementById('qword-input-sink');
    const sr = sink.getBoundingClientRect();
    const caretEl = (function(){ const list = document.querySelectorAll('.qword-caret'); for (const c of list) if (c.style.display !== 'none') return c; return null; })();
    const cr = caretEl ? caretEl.getBoundingClientRect() : null;
    const A = window.QWordEngine;
    const r = A.caretRect(A.caret.para, A.caret.cp);
    const cs = getComputedStyle(sink);
    return JSON.stringify({ sink: { left: sr.left, top: sr.top, w: sr.width, h: sr.height }, caret: cr ? { left: cr.left, top: cr.top, h: cr.height } : null, opacity: cs.opacity, pe: cs.pointerEvents, rect: r });
  })()`));
  const dx = g3.caret ? Math.abs(g3.sink.left - g3.caret.left) : Infinity;
  const dy = g3.caret ? Math.abs(g3.sink.top - g3.caret.top) : Infinity;
  const offscreenConst = Math.abs(g3.sink.left + 9999) < 1 && Math.abs(g3.sink.top + 9999) < 1;
  if (!g3.caret) fail(`${tag}/L3 候選窗錨點鎖`, '找不到可見的 .qword-caret，無法比對尺');
  else if (offscreenConst) fail(`${tag}/L3 候選窗錨點鎖`, `inputSink 仍在 -9999 離屏（left=${g3.sink.left}, top=${g3.sink.top}）→ macOS 候選窗开到螢幕外`);
  else if (dx > 2 || dy > 2) fail(`${tag}/L3 候選窗錨點鎖`, `sink(${Math.round(g3.sink.left)},${Math.round(g3.sink.top)}) 與畫面游標(${Math.round(g3.caret.left)},${Math.round(g3.caret.top)}) 差 dx=${dx.toFixed(1)} dy=${dy.toFixed(1)} >2px`);
  else if (Number(g3.opacity) > 0) fail(`${tag}/L3 候選窗錨點鎖`, `sink 未維持不可見姿態 opacity=${g3.opacity}`);
  else pass(`${tag}/L3 候選窗錨點鎖`, `sink 錨在游標 dx=${dx.toFixed(1)} dy=${dy.toFixed(1)}px（≤2px）、opacity=${g3.opacity}、pointer-events=${g3.pe}`);

  // ---------- L4 合成回顯鎖 ----------
  const beforeP4 = await evalIn(`window.QWordEngine.paragraphText(${P})||''`);
  await clickIntoCaret(P, Array.from(beforeP4).length);
  await clearTrace();
  await cdp.send('Input.imeSetComposition', { text: 'zhong', selectionStart: 5, selectionEnd: 5 });
  await sleep(340);
  const mid4 = JSON.parse(await evalIn(`(() => {
    const ov = document.querySelector('.qword-composition-overlay');
    const A = window.QWordEngine;
    const all = [];
    for (let p = 0; p < A.paragraphCount(); p++) all.push(A.paragraphText(p) || '');
    return JSON.stringify({ ov: ov ? { text: ov.textContent, visible: getComputedStyle(ov).display !== 'none' && getComputedStyle(ov).visibility !== 'hidden', cls: ov.className } : null,
      engineHas: all.some(t => t.includes('zhong')), bodyHas: (document.querySelector('.qword-page-container').innerText || '').includes('zhong'),
      ovInCaretLayer: !!(ov && ov.closest('.qword-caret-layer')) });
  })()`));
  const s4 = JSON.parse(await dump());
  const winOpen4 = s4.trace.some((x) => x.t === 'compositionupdate');
  if (!winOpen4) fail(`${tag}/L4 合成回顯鎖`, 'compositionupdate 缺席——事件序未成立');
  else if (!mid4.ov || !mid4.ov.visible) fail(`${tag}/L4 合成回顯鎖`, '合成期 DOM 無可見的合成回顯 overlay（用戶「打了沒反應」根因）');
  else if (!String(mid4.ov.text).includes('zhong')) fail(`${tag}/L4 合成回顯鎖`, `overlay 文字=${JSON.stringify(mid4.ov.text)} 未含拼音`);
  else if (mid4.engineHas) fail(`${tag}/L4 合成回顯鎖`, '拼音寫進引擎正文（合成期零引擎寫入鐵律破）');
  else if (mid4.bodyHas) fail(`${tag}/L4 合成回顯鎖`, '拼音洩漏進正文 DOM');
  else if (!mid4.ovInCaretLayer) fail(`${tag}/L4 合成回顯鎖`, `overlay 不在 caret 層（class=${mid4.ov.cls}）——與游標不同尺`);
  else {
    await cdp.send('Input.insertText', { text: '中文' });
    await sleep(520);
    const after4 = JSON.parse(await evalIn(`(() => {
      const ov = document.querySelector('.qword-composition-overlay');
      const A = window.QWordEngine;
      return JSON.stringify({ ovGone: !ov || getComputedStyle(ov).display === 'none' || String(ov.textContent) === '', para: A.paragraphText(${P}) || '', hasZh: (A.paragraphText(${P})||'').includes('中文') });
    })()`));
    const noLeak = !String(after4.para).includes('zhong');
    if (!after4.hasZh || !noLeak) fail(`${tag}/L4 合成回顯鎖`, `compositionend 後正文=${JSON.stringify(after4.para.slice(-10))}（定稿未入或拼音殘留）`);
    else if (!after4.ovGone) fail(`${tag}/L4 合成回顯鎖`, 'compositionend 後合成 overlay 未清除');
    else pass(`${tag}/L4 合成回顯鎖`, `合成期回顯「zhong」（零引擎寫入）→ 定稿入文「中文」＋overlay 清除`);
  }

  // ---------- L5 入文點快照鎖 ----------
  // 合成窗開著時以另一途徑（真點擊另一段）移動 caret → 定稿仍落 compositionstart 快照點。
  const Q = await evalIn(`(() => { const A = window.QWordEngine; for (let p = ${P} + 1; p < A.paragraphCount(); p++) { if (Array.from(A.paragraphText(p)||'').length > 3) return p; } return -1; })()`);
  if (Q < 0) {
    fail(`${tag}/L5 入文點快照鎖`, '找不到第二個可跳躍的正文段');
  } else {
    const len5 = await evalIn(`Array.from(window.QWordEngine.paragraphText(${P})||'').length`);
    const snapP5 = await evalIn(`window.QWordEngine.paragraphText(${P})||''`);
    const snapQ5 = await evalIn(`window.QWordEngine.paragraphText(${Q})||''`);
    await clickIntoCaret(P, len5);
    await clearTrace();
    await cdp.send('Input.imeSetComposition', { text: 'kecheng', selectionStart: 7, selectionEnd: 7 });
    await sleep(300);
    const winOpen5 = JSON.parse(await dump()).trace.some((x) => x.t === 'compositionstart');
    await clickIntoCaret(Q, 2);                          // 合成窗開著：另一途徑移動 caret
    await sleep(200);
    await cdp.send('Input.insertText', { text: '課程' });
    await sleep(520);
    const s5 = JSON.parse(await dump());
    const pText = s5.paras[P] || '';
    const qText = s5.paras[Q] || '';
    const landedSnapshot = pText.length === Array.from(snapP5).length + 2 && pText.includes('課程');
    const drifted = qText !== snapQ5;
    if (!winOpen5) fail(`${tag}/L5 入文點快照鎖`, '合成窗未開——測試無效');
    else if (drifted) fail(`${tag}/L5 入文點快照鎖`, `定稿落進合成期被移動後的段 ${Q}（=${JSON.stringify(qText.slice(-8))}）而非 compositionstart 快照段 ${P}`);
    else if (!landedSnapshot) fail(`${tag}/L5 入文點快照鎖`, `段 ${P} 未見快照點入文：${JSON.stringify(pText.slice(-10))}`);
    else pass(`${tag}/L5 入文點快照鎖`, '合成期 caret 被外部移動 → 定稿仍落 compositionstart 快照點');
  }
}

for (const dpr of DPRS) await runRound(dpr);

// ---------- 零未捕獲例外前置燈（殼碼 ReferenceError 會被瀏覽器吞掉→整臂靜默不執行＝假綠）----------
if (pageErrors.length) fail('全域 零未捕獲 JS 例外', `${pageErrors.length} 筆：${pageErrors.slice(0, 2).join(' | ')}`);
else pass('全域 零未捕獲 JS 例外', 'pageerror 計數 0');

fs.mkdirSync(path.dirname(OUT), { recursive: true });
fs.writeFileSync(OUT, JSON.stringify({ at: new Date().toISOString(), baseUrl: BASE_URL, dprs: DPRS, buggy: BUGGY || null, results }, null, 2));

try { process.kill(-httpChild.pid, 'SIGKILL'); } catch {}
try { ws.close(); } catch {}
chromeChild.kill('SIGKILL');

const fails = results.filter((r) => !r.ok);
console.log(`\n判定：${fails.length === 0 ? 'PASS' : 'FAIL'}（${fails.length} 項紅）｜證據 ${OUT}`);
if (BUGGY) console.log(`（回塞態 U6_IME_BUGGY=${BUGGY}：對應燈按口徑應紅——紅數即鑑別力）`);
process.exit(fails.length === 0 ? 0 : 1);
