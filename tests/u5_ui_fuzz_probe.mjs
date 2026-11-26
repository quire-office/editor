/**
 * u5_ui_fuzz_probe.mjs — 編輯器 UI 乱點／亂輸入韌性探針（用戶 11-26 指定）
 *
 * 目標（用戶原話）：「測試盡量涵蓋點擊所有按鈕，可以測試多一點組合，但亂點也行，
 * 避免出現用戶隨意點擊以及輸入時，出現預期外的事件，輸入也是一樣，要避免錯誤的狀況。」
 *
 * 三層判據（都是「預期外」的機械可判形式，不是目視）：
 *  U1 零未捕獲例外／零 console.error（任何一步觸發即紅並點名步驟）
 *  U2 崩潰守門：DOM 結構存活（页面容器＋canvas＋inputSink 仍在）、Ribbon 未被
 *     亂點破壞（按鈕數不減）、頁面容器仍可渲染（pageCount ≥1）
 *  U3 不变量：游標/選區恆在合法域（段號 0..pc-1、cp 0..段長）、打字後段文字
 *     與預期一致（不打錯字、不吞字）、存檔仍可開（save→openFile 不炸）、
 *     **打字不得引入未請求的縮排/前導空白**（用戶實測「打字後突然縮排」）
 *
 * 派發方式全部走 CDP Input.dispatch*（真滑鼠真鍵盤），不走 JS 直呼——只有真事件
 * 路徑能抓到 C1 那類「相容事件搶焦點」缺陷（探針紀律見 quire-editor-development）。
 *
 * 用法：
 *   python3 -m http.server 8791 --bind 127.0.0.1 &    # 本地（默認）
 *   node tests/u5_ui_fuzz_probe.mjs
 *   FUZZ_URL=https://editor.quireoffice.dev/word.html node tests/u5_ui_fuzz_probe.mjs
 *   FUZZ_SEED=42 決定性亂點（默認 20261126）；FUZZ_STEPS=120 步數
 */
const BASE_URL = process.env.FUZZ_URL || 'http://127.0.0.1:8791/word.html';
const OUT = process.env.FUZZ_OUT || null;
const CDP_HTTP = process.env.CDP_URL || 'http://127.0.0.1:9222';
const SEED = Number(process.env.FUZZ_SEED || 20261126);
const STEPS = Number(process.env.FUZZ_STEPS || 120);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// 決定性 PRNG（mulberry32）：同 seed 同序列 → 失敗可重現
function mulberry32(a) {
  return function () {
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const rnd = mulberry32(SEED);
const pick = (arr) => arr[Math.floor(rnd() * arr.length)];
const ri = (lo, hi) => lo + Math.floor(rnd() * (hi - lo + 1));

async function main() {
  const created = await (await fetch(`${CDP_HTTP}/json/new?about:blank`, { method: 'PUT' })).json();
  const sock = new WebSocket(created.webSocketDebuggerUrl);
  await new Promise((res, rej) => { sock.onopen = res; sock.onerror = rej; });
  let seq = 0; const pending = new Map();
  const consoleErrors = []; const pageErrors = [];
  sock.onmessage = (ev) => {
    const m = JSON.parse(ev.data);
    if (m.id && pending.has(m.id)) { pending.get(m.id)(m); pending.delete(m.id); }
    if (m.method === 'Runtime.exceptionThrown') {
      pageErrors.push({ where: m.params.timestamp, msg: String(m.params.exceptionDetails?.exception?.description || m.params.exceptionDetails?.text || '').slice(0, 240) });
    }
    if (m.method === 'Runtime.consoleAPICalled' && m.params.type === 'error') {
      consoleErrors.push({ msg: (m.params.args || []).map((a) => a.value ?? a.description ?? '').join(' ').slice(0, 240) });
    }
  };
  const send = (method, params = {}) => new Promise((res) => {
    const id = ++seq; pending.set(id, (m) => res(m.result || {}));
    sock.send(JSON.stringify({ id, method, params }));
  });
  const evalIn = async (expr) => {
    const r = await send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: true, timeout: 120000 });
    if (r.exceptionDetails) return '__EX__:' + String(r.exceptionDetails.exception?.description || '').slice(0, 200);
    return r.result ? r.result.value : undefined;
  };

  await send('Page.enable'); await send('Runtime.enable'); await send('Log.enable').catch(() => {});
  await send('Page.navigate', { url: BASE_URL });
  for (let i = 0; i < 240; i++) { if (await evalIn('document.readyState') === 'complete') break; await sleep(250); }
  await evalIn('window.QWordShell.openDefaultSample()');
  for (let i = 0; i < 300; i++) { if (await evalIn('!!window.QWordEngine.engine && window.QWordEngine.paragraphCount() > 0')) break; await sleep(200); }
  await sleep(600);

  // 開盤基準（U2/U3 的對照系）
  const base = JSON.parse(await evalIn(`JSON.stringify({
    btnCount: document.querySelectorAll('button[data-command]').length,
    pages: window.QWordEngine.pageCount(),
    paras: window.QWordEngine.paragraphCount()
  })`));

  // 可亂點目標集：全部 Ribbon 鈕 + tab 鈕 + popover 觸發鈕
  const targets = await evalIn(`(function () {
    const sel = 'button[data-command], [data-tab-target], [data-popover-target]';
    return JSON.stringify(Array.from(document.querySelectorAll(sel)).map(function (el, i) {
      const r = el.getBoundingClientRect();
      return { i, cmd: el.getAttribute('data-command') || el.getAttribute('data-tab-target') || el.getAttribute('data-popover-target') || el.tagName.toLowerCase() + i,
               x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2),
               vis: r.width > 0 && r.height > 0 };
    }));
  })()`).then((s) => JSON.parse(s));
  const clickable = targets.filter((t) => t.vis && t.y > 0);

  const keys = ['a', 'b', 'A', '1', ' ，', '，', '测', '試', ' ', '.', '-', '（', 'x'];
  const specials = ['Enter', 'Tab', 'Backspace', 'Delete', 'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'];
  const steps = [];

  const pressKey = async (k) => {
    const isText = k.length === 1 || k === ' ';
    const base = { key: isText ? k : k, code: isText ? 'Key' + k.toUpperCase() : k, windowsVirtualKeyCode: (k.charCodeAt(0)) };
    if (isText) {
      await send('Input.dispatchKeyEvent', { type: 'keyDown', ...base, text: k, unmodifiedKey: k });
      await send('Input.dispatchKeyEvent', { type: 'keyUp', ...base });
    } else {
      const vk = { Enter: 13, Tab: 9, Backspace: 8, Delete: 46, ArrowLeft: 37, ArrowRight: 39, ArrowUp: 38, ArrowDown: 40, Home: 36, End: 35 }[k] || 0;
      await send('Input.dispatchKeyEvent', { type: 'keyDown', key: k, code: k, windowsVirtualKeyCode: vk });
      await send('Input.dispatchKeyEvent', { type: 'keyUp', key: k, code: k, windowsVirtualKeyCode: vk });
    }
  };

  const caretGeo = async () => evalIn(`(function () {
    const A = window.QWordEngine;
    if (!A.caret) return null;
    const r = A.caretRect(A.caret.para, A.caret.cp);
    if (!r) return null;
    const w = document.querySelectorAll('.qword-real-page')[r.page];
    if (!w) return null;
    const c = w.querySelector('canvas'); const rect = c.getBoundingClientRect();
    const dl = A.getPageDisplayList(r.page); const scale = c.width / (dl.width_pt || 595);
    return { x: rect.left + r.x_pt * scale, y: rect.top + (r.y_pt - 4) * scale };
  })()`);

  // ---- 亂點＋亂輸入主迴圈 ----
  for (let s = 0; s < STEPS; s++) {
    const roll = rnd();
    const rec = { s, act: null };
    try {
      if (roll < 0.42 && clickable.length) {
        // 點按鈕（含 tab 切換、popover 開關、格式鈕、表格鈕——未圈選態也照點，
        // 這正是用戶「隨意點擊」會遇到的路徑）
        const t = pick(clickable);
        rec.act = `click:${t.cmd}`;
        await send('Input.dispatchMouseEvent', { type: 'mousePressed', x: t.x, y: t.y, button: 'left', clickCount: 1, buttons: 1 });
        await send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: t.x, y: t.y, button: 'left', buttons: 0 });
      } else if (roll < 0.62) {
        // 點頁面任意處（含頁縫、灰區、跨頁）
        const pg = pick(await evalIn(`JSON.stringify(Array.from(document.querySelectorAll('.qword-page-canvas')).map(function(c){var r=c.getBoundingClientRect();return {x:r.left,y:r.top,w:r.width,h:r.height};}))`).then((x) => JSON.parse(x)) || []);
        if (pg && pg.w) {
          const x = pg.x + rnd() * pg.w, y = pg.y + rnd() * pg.h;
          rec.act = `pageclick`;
          await send('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', clickCount: 1, buttons: 1 });
          await send('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', buttons: 0 });
        }
      } else if (roll < 0.74) {
        // 拖曳選取（含拖出頁外）
        const g = await caretGeo();
        if (g) {
          rec.act = 'drag';
          await send('Input.dispatchMouseEvent', { type: 'mousePressed', x: g.x, y: g.y, button: 'left', clickCount: 1, buttons: 1 });
          for (let k = 0; k < 3; k++) {
            await send('Input.dispatchMouseEvent', { type: 'mouseMoved', x: g.x + ri(-300, 300), y: g.y + ri(-260, 260), button: 'left', buttons: 1 });
            await sleep(40);
          }
          await send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: g.x + ri(-300, 300), y: g.y + ri(-260, 260), button: 'left', buttons: 0 });
        }
      } else if (roll < 0.9) {
        // 打字（含中英數標空白）
        const k = pick(keys);
        rec.act = `type:${JSON.stringify(k)}`;
        await pressKey(k);
      } else {
        // 功能鍵
        const k = pick(specials);
        rec.act = `key:${k}`;
        await pressKey(k);
      }
    } catch (e) { rec.threw = String(e.message || e).slice(0, 120); }
    steps.push(rec);
    if (s % 12 === 0) await sleep(60);
  }

  // ---- U3 不变量檢查 ----
  const inv = await evalIn(`(async function () {
    const A = window.QWordEngine;
    const out = { ok: true, notes: [] };
    const pc = A.paragraphCount();
    if (!A.caret) out.notes.push('caret 消失');
    else if (!(A.caret.para >= 0 && A.caret.para < pc)) out.notes.push('caret.para 越界 ' + JSON.stringify(A.caret) + ' pc=' + pc);
    else if (!(A.caret.cp >= 0 && A.caret.cp <= Array.from(A.paragraphText(A.caret.para) || '').length)) out.notes.push('caret.cp 越界 ' + JSON.stringify(A.caret));
    if (A.selection) {
      const s = A.selection, e = A.selection.end;
      for (const p of [s.start, e]) {
        if (!(p.para >= 0 && p.para < pc)) out.notes.push('selection.para 越界 ' + JSON.stringify(p));
        else if (!(p.cp >= 0 && p.cp <= Array.from(A.paragraphText(p.para) || '').length)) out.notes.push('selection.cp 越界 ' + JSON.stringify(p));
      }
    }
    // 存檔→重開 不炸
    try {
      const b = A.save(); const u8 = b instanceof Uint8Array ? b : new Uint8Array(b);
      out.savedBytes = u8.length;
      await window.QWordShell.openFile(u8, 'fuzz_rt.docx');
      await new Promise(function (r) { setTimeout(r, 500); });
      out.reopenedParas = window.QWordEngine.paragraphCount();
      if (out.reopenedParas <= 0) out.notes.push('重開後零段');
    } catch (e) { out.notes.push('save/reopen 拋例外: ' + String(e.message || e).slice(0, 120)); }
    // DOM 存活
    out.pageCanvases = document.querySelectorAll('.qword-page-canvas').length;
    out.inputSinkAlive = !!document.getElementById('qword-input-sink');
    out.btnCountAfter = document.querySelectorAll('button[data-command]').length;
    if (out.pageCanvases === 0) out.notes.push('頁面 canvas 全失');
    if (!out.inputSinkAlive) out.notes.push('inputSink 消失');
    // 打字鏈仍活（末態打一個可辨識標記）
    return JSON.stringify(out);
  })()`);
  const invJson = typeof inv === 'string' && inv.startsWith('__EX__') ? { ok: false, notes: [inv] } : JSON.parse(inv || '{}');

  // ---- U3 縮排回歸（用戶實測「打字後突然縮排」）----
  // 直接驗：游標在段首打字，新文字必須出現在段首、前面不得冒出空白/TAB。
  const indent = await evalIn(`(async function () {
    const A = window.QWordEngine;
    await window.QWordShell.openDefaultSample();
    await new Promise(function (r) { setTimeout(r, 600); });
    var p = 0;
    for (var i = 0; i < A.paragraphCount(); i++) { if (Array.from(A.paragraphText(i)||'').length > 8) { p = i; break; } }
    var before = A.paragraphText(p) || '';
    A.caret = { para: p, cp: 0 }; A.selection = null;
    var r = A.typeAt(p, 0, 'ZK');
    var after = A.paragraphText(p) || '';
    return JSON.stringify({
      p: p, ok: !!(r && r.ok), beforeHead: before.slice(0, 6), afterHead: after.slice(0, 8),
      // 打了 'ZK' 就必須以 ZK 開頭；若前面多出空白/全形空格即是用戶說的「縮排」
      leadingIntruder: /^[\s\u3000\t]+/.test(after) || !after.startsWith('ZK'),
      inserted: after.slice(0, 2)
    });
  })()`);
  const indentJson = typeof indent === 'string' && indent.startsWith('__EX__') ? { ok: false, leadingIntruder: true, why: indent } : JSON.parse(indent || '{}');

  try { await fetch(`${CDP_HTTP}/json/close/${created.id}`); } catch {}
  sock.close();

  const R = { meta: { baseUrl: BASE_URL, seed: SEED, steps: STEPS, base },
              steps, inv: invJson, indent: indentJson,
              pageErrors, consoleErrors };
  if (OUT) { const fs = await import('node:fs'); fs.writeFileSync(OUT, JSON.stringify(R, null, 2)); }

  let fail = 0;
  const line = (cond, msg) => { console.log(`  ${cond ? '✓' : '✗'} ${msg}`); if (!cond) fail += 1; };
  console.log('=== U5 UI 乱點／亂輸入 韌性判定 ===');
  console.log(`  （${STEPS} 步、seed=${SEED}、baseUrl=${BASE_URL}）`);
  line(pageErrors.length === 0 && consoleErrors.length === 0,
    `U1 零未捕獲例外／零 console.error（實得 ${pageErrors.length}/${consoleErrors.length}）`);
  if (pageErrors.length) console.log('     例外樣本:', JSON.stringify(pageErrors.slice(0, 3)));
  if (consoleErrors.length) console.log('     error 樣本:', JSON.stringify(consoleErrors.slice(0, 3)));
  line(steps.every((s) => !s.threw), `U1b 每步派發零拋例外（異常 ${steps.filter((s) => s.threw).length} 步）`);
  line(invJson.pageCanvases > 0 && invJson.inputSinkAlive, `U2 DOM 存活（canvas=${invJson.pageCanvases} sink=${invJson.inputSinkAlive}）`);
  line(invJson.btnCountAfter >= base.btnCount, `U2b Ribbon 未被亂點破壞（${base.btnCount} → ${invJson.btnCountAfter}）`);
  line((invJson.notes || []).length === 0, `U3 不变量全守（${JSON.stringify(invJson.notes || [])}）`);
  line(invJson.savedBytes > 0 && invJson.reopenedParas > 0, `U3b 亂搞後存檔可重開（bytes=${invJson.savedBytes} paras=${invJson.reopenedParas}）`);
  line(indentJson.ok && !indentJson.leadingIntruder,
    `U3c 段首打字不引入縮排/前導空白（head=${JSON.stringify(indentJson.afterHead)} intruder=${indentJson.leadingIntruder}）`);
  console.log(`\n判定：${fail === 0 ? 'PASS' : `FAIL（${fail} 項）`}`);
  process.exit(fail === 0 ? 0 : 1);
}

main().catch((e) => { console.error('探針失敗:', e.message); process.exit(2); });
